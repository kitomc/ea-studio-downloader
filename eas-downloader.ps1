# eas-downloader.ps1
# Descarga estrategias EAS con soporte para collections grandes (300+).
# MQ5 + MQ4 + CSV + log.txt todo en la misma carpeta.

Add-Type -AssemblyName System.Windows.Forms

Write-Host "[INFO] Selecciona el archivo JSON de la collection..." -ForegroundColor Cyan
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "JSON Files|*.json|All Files|*.*"
$openFileDialog.Title = "Selecciona la Collection de EAS"
if (Test-Path "C:\Users\kitom\Downloads\EA Studio Downloader") {
    $openFileDialog.InitialDirectory = "C:\Users\kitom\Downloads\EA Studio Downloader"
} else {
    $openFileDialog.InitialDirectory = $PSScriptRoot
}
if ($openFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "[ERROR] No se selecciono ningun archivo." -ForegroundColor Red
    exit 1
}
$CollectionFile = $openFileDialog.FileName
Write-Host "[INFO] Collection: $CollectionFile" -ForegroundColor Green

try { python --version | Out-Null } catch {
    Write-Host "[ERROR] Python no instalado" -ForegroundColor Red
    exit 1
}

python -c "import playwright" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[INFO] Instalando playwright..." -ForegroundColor Yellow
    pip install playwright -q 2>$null
    playwright install chromium 2>$null
}

$edgeUserData = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
$downloadBase = "C:\Users\kitom\Downloads\EA Studio Downloader\downloads"
New-Item -ItemType Directory -Path $downloadBase -Force | Out-Null

$collectionJson = Get-Content $CollectionFile -Raw
$strategyCount = ($collectionJson | ConvertFrom-Json).Count
Write-Host "[INFO] Collection tiene $strategyCount estrategias" -ForegroundColor Cyan
Write-Host "[INFO] Descargas en: $downloadBase" -ForegroundColor Cyan

$logFile = Join-Path $downloadBase "log.txt"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $logFile -Value "[$timestamp] === INICIO DESCARGA ==="
Add-Content -Path $logFile -Value "[$timestamp] Collection: $CollectionFile"
Add-Content -Path $logFile -Value "[$timestamp] Estrategias esperadas: $strategyCount"
Add-Content -Path $logFile -Value "[$timestamp] Directorio descargas: $downloadBase"

$pyScript = @"
import asyncio, json, os, sys, time
from pathlib import Path
from playwright.async_api import async_playwright

EDGE_USER_DATA = os.path.join(os.environ["LOCALAPPDATA"], "Microsoft", "Edge", "User Data")
COLLECTION_FILE = r"$CollectionFile"
DOWNLOAD_DIR = Path(r"$downloadBase")
LOG_FILE = DOWNLOAD_DIR / "log.txt"
EXPECTED_TOTAL = $strategyCount

def log(msg):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

async def main():
    log("Iniciando navegador Edge con perfil Kito...")
    
    import subprocess
    subprocess.run(["taskkill", "/F", "/IM", "msedge.exe"], capture_output=True)
    await asyncio.sleep(2)
    
    async with async_playwright() as p:
        browser = await p.chromium.launch_persistent_context(
            user_data_dir=EDGE_USER_DATA,
            headless=False,
            channel="msedge",
            accept_downloads=True,
            args=[
                "--profile-directory=Default",
                "--disable-blink-features=AutomationControlled",
            ],
            no_viewport=True,
        )
        
        page = browser.pages[0] if browser.pages else await browser.new_page()
        
        log("Navegando a expert-advisor-studio.com...")
        await page.goto("https://expert-advisor-studio.com/#", wait_until="networkidle", timeout=30000)
        await page.wait_for_timeout(2000)
        
        local_storage = await page.evaluate("""() => {
            const data = {};
            for (let i = 0; i < localStorage.length; i++) {
                const key = localStorage.key(i);
                data[key] = localStorage.getItem(key);
            }
            return data;
        }""")
        has_user = "eaStudio-user" in local_storage
        has_premium = any("Premium" in str(v) for v in local_storage.values())
        log(f"Sesion activa: {has_user}, Premium: {has_premium}")
        
        if not has_user:
            log("ERROR: No hay sesion activa. Logueate en EAS y reintenta.")
            await browser.close()
            return
        
        # Ir a Collection
        log("Navegando a Collection...")
        await page.click("#eas-navbar-collection-link")
        await page.wait_for_load_state("networkidle", timeout=15000)
        await page.wait_for_timeout(1500)
        
        # Subir collection
        log(f"Subiendo collection ({EXPECTED_TOTAL} estrategias)...")
        upload_button = await page.query_selector("#upload-button")
        if not upload_button:
            upload_button = await page.query_selector("div:nth-of-type(3) span.eas-button-text")
        
        async with page.expect_file_chooser(timeout=10000) as fc_info:
            await upload_button.click()
        file_chooser = await fc_info.value
        await file_chooser.set_files(COLLECTION_FILE)
        log("Collection subida!")
        
        await page.wait_for_selector('[id^="collection-record-"]', timeout=15000)
        await page.wait_for_timeout(2000)
        
        initial_count = len(await page.query_selector_all('[id^="collection-record-"]'))
        log(f"Registros visibles despues de upload: {initial_count}")
        
        # Screenshot para debug
        await page.screenshot(path=str(DOWNLOAD_DIR / "debug_after_upload.png"))
        log("Screenshot guardado: debug_after_upload.png")
        
        # Inspeccionar paginacion
        pagination_html = await page.evaluate("""() => {
            const pagination = document.querySelector('.pagination, [class*="pagination"], [class*="pager"]');
            if (pagination) return pagination.outerHTML.substring(0, 1000);
            return 'NO_PAGINATION_FOUND';
        }""")
        log(f"HTML paginacion: {pagination_html[:200]}")
        
        seen_ids = set()
        total_ok = 0
        total_mq4 = 0
        total_csv = 0
        total_err = 0
        errors_log = []
        consecutive_same = 0
        start_time = time.time()
        
        # BUCLE PRINCIPAL: Ctrl+ArrowRight para navegar
        # Flujo por estrategia: MQ4 -> MQ5 -> CSV -> Editor -> Ctrl+ArrowRight
        log(f"Descargando {EXPECTED_TOTAL} estrategias...")
        
        # Click en primera estrategia
        first_record = await page.query_selector('[id^="collection-record-"]')
        if not first_record:
            log("ERROR: No hay registros en la collection.")
            await browser.close()
            return
        
        await first_record.click()
        try:
            await page.wait_for_selector("#editor-toolbar-export", timeout=10000)
            await page.wait_for_timeout(500)
        except:
            log("ERROR: Editor no cargo.")
            await browser.close()
            return
        
        # Esperar a que la strategy card cargue y hacer click en el canvas para focus
        await page.wait_for_timeout(500)
        
        strategy_num = 0
        last_id = None
        
        while strategy_num < EXPECTED_TOTAL:
            strategy_num += 1
            
            # Leer ID de la estrategia actual
            current_id = None
            try:
                id_el = await page.query_selector("#navbar-strategy-id")
                if id_el:
                    current_id = await id_el.inner_text()
            except:
                pass
            
            # Detectar fin de collection (mismo ID repetido)
            if current_id and current_id == last_id:
                consecutive_same += 1
                if consecutive_same >= 5:
                    log(f"Mismo ID '{current_id}' 5 veces seguidas. Fin de la collection en {strategy_num - 1}.")
                    strategy_num -= 1
                    break
            else:
                consecutive_same = 0
            last_id = current_id
            
            elapsed = time.time() - start_time
            rate = strategy_num / elapsed if elapsed > 0 else 0
            eta = (EXPECTED_TOTAL - strategy_num) / rate if rate > 0 else 0
            log(f"[{strategy_num}/{EXPECTED_TOTAL}] {current_id} (~{eta:.0f}s restantes)")
            
            # ---- MQ4 ----
            mq5_stem = None
            try:
                await page.click("#editor-toolbar-export")
                await page.wait_for_selector("#editor-toolbar-export-ea4", timeout=3000)
                await page.wait_for_timeout(200)
                async with page.expect_download(timeout=20000) as dl:
                    await page.click("#editor-toolbar-export-ea4")
                download = await dl.value
                mq4_name = download.suggested_filename
                if not mq4_name.endswith(".mq4"):
                    mq4_name += ".mq4"
                mq5_stem = mq4_name.replace(".mq4", "")
                await download.save_as(DOWNLOAD_DIR / mq4_name)
                log(f"  [OK] MQ4: {mq4_name}")
                total_mq4 += 1
            except Exception as e:
                log(f"  [ERR] MQ4: {e}")
                total_err += 1
                errors_log.append(f"{strategy_num}: MQ4 - {e}")
            
            # ---- MQ5 ----
            if mq5_stem:
                try:
                    await page.click("#editor-toolbar-export")
                    await page.wait_for_selector("#editor-toolbar-export-ea5", timeout=3000)
                    await page.wait_for_timeout(200)
                    async with page.expect_download(timeout=20000) as dl:
                        await page.click("#editor-toolbar-export-ea5")
                    download = await dl.value
                    mq5_name = download.suggested_filename
                    if not mq5_name.endswith(".mq5"):
                        mq5_name += ".mq5"
                    mq5_stem_clean = Path(mq5_name).stem
                    
                    if mq5_stem_clean in seen_ids:
                        log(f"  [SKIP] Duplicado: {mq5_stem_clean}")
                        await download.save_as(DOWNLOAD_DIR / mq5_name)
                        strategy_num -= 1
                    else:
                        seen_ids.add(mq5_stem_clean)
                        await download.save_as(DOWNLOAD_DIR / mq5_name)
                        log(f"  [OK] MQ5: {mq5_name}")
                        total_ok += 1
                except Exception as e:
                    log(f"  [ERR] MQ5: {e}")
                    total_err += 1
                    errors_log.append(f"{strategy_num}: MQ5 - {e}")
            
            # ---- CSV (Report -> Journal -> Export) ----
            if mq5_stem:
                try:
                    await page.click("#eas-navbar-report-link")
                    await page.wait_for_timeout(800)
                    journal_tab = await page.query_selector("#journal-tab")
                    if journal_tab:
                        await journal_tab.click()
                        await page.wait_for_timeout(600)
                    export_btn = await page.query_selector("#report-journal-export")
                    if export_btn:
                        async with page.expect_download(timeout=20000) as csv_dl:
                            await export_btn.click()
                        csv_download = await csv_dl.value
                        stem_for_csv = mq5_stem if not mq5_stem.startswith("EA Studio") else mq5_stem
                        csv_name = f"{stem_for_csv}.csv"
                        await csv_download.save_as(DOWNLOAD_DIR / csv_name)
                        log(f"  [OK] CSV: {csv_name}")
                        total_csv += 1
                except Exception as e:
                    log(f"  [--] CSV: {e}")
            
            # ---- Volver al Editor ----
            try:
                editor_link = await page.query_selector("#eas-navbar-editor-link")
                if editor_link:
                    await editor_link.click()
                    await page.wait_for_timeout(500)
            except:
                pass
            
            # ---- Ctrl+ArrowRight para siguiente estrategia ----
            if strategy_num < EXPECTED_TOTAL:
                try:
                    # Click en tabla para asegurar focus
                    table_cell = await page.query_selector("#backtest-output-table td")
                    if table_cell:
                        await table_cell.click()
                        await page.wait_for_timeout(200)
                    
                    await page.keyboard.press("Control+ArrowRight")
                    await page.wait_for_timeout(1500)
                    try:
                        await page.wait_for_selector("#editor-toolbar-export", timeout=8000)
                    except:
                        log("  [WARN] Editor no recargo, reintentando Ctrl+ArrowRight...")
                        await page.keyboard.press("Control+ArrowRight")
                        await page.wait_for_timeout(1500)
                        await page.wait_for_selector("#editor-toolbar-export", timeout=8000)
                except Exception as e:
                    log(f"  [ERR] No se pudo navegar a siguiente: {e}")
                    break
        
        elapsed = time.time() - start_time
        await browser.close()
        
        # Resumen en log
        mq5_files = list(DOWNLOAD_DIR.glob("*.mq5"))
        mq4_files = list(DOWNLOAD_DIR.glob("*.mq4"))
        csv_files = list(DOWNLOAD_DIR.glob("*.csv"))
        
        log("=" * 55)
        log(f"RESUMEN FINAL")
        log(f"  Estrategias esperadas  : {EXPECTED_TOTAL}")
        log(f"  Unicos descargados    : {len(seen_ids)}")
        log(f"  Exitosas              : {total_ok}")
        log(f"  Errores               : {total_err}")
        log(f"  MQ5 descargados       : {len(mq5_files)}")
        log(f"  MQ4 descargados       : {len(mq4_files)}")
        log(f"  CSV descargados       : {len(csv_files)}")
        log(f"  Tiempo                : {elapsed:.1f}s ({elapsed/60:.1f} min)")
        if len(seen_ids) > 0:
            log(f"  Velocidad            : {len(seen_ids)/elapsed*60:.1f} estr/min")
        if len(seen_ids) < EXPECTED_TOTAL:
            log(f"  FALTAN               : {EXPECTED_TOTAL - len(seen_ids)} estrategias!!")
        if errors_log:
            log("ERRORES:")
            for err in errors_log:
                log(f"  - {err}")
        log("=" * 55)

asyncio.run(main())
"@

$pyScript | Out-File -FilePath "eas_downloader.py" -Encoding UTF8
Write-Host "[INFO] Iniciando descarga de $strategyCount estrategias..." -ForegroundColor Cyan

python eas_downloader.py

# Validar
Write-Host "`n[VALIDACION] Verificando resultados..." -ForegroundColor Cyan

$mq5Files = Get-ChildItem "$downloadBase\*.mq5" -ErrorAction SilentlyContinue
$mq4Files = Get-ChildItem "$downloadBase\*.mq4" -ErrorAction SilentlyContinue
$csvFiles  = Get-ChildItem "$downloadBase\*.csv"  -ErrorAction SilentlyContinue

$mq5Count = if ($mq5Files) { $mq5Files.Count } else { 0 }
$mq4Count = if ($mq4Files) { $mq4Files.Count } else { 0 }
$csvCount = if ($csvFiles) { $csvFiles.Count } else { 0 }

$pct = if ($strategyCount -gt 0) { [math]::Round($mq5Count / $strategyCount * 100, 1) } else { 0 }
Write-Host "[RESULTADO] MQ5: $mq5Count | MQ4: $mq4Count | CSV: $csvCount de $strategyCount ($pct%)" -ForegroundColor $(if ($pct -ge 100) {"Green"} else {"Yellow"})

if (Test-Path (Join-Path $downloadBase "log.txt")) {
    Write-Host "`n[LOG] Ultimas 30 lineas de log.txt:" -ForegroundColor Cyan
    Get-Content (Join-Path $downloadBase "log.txt") -Tail 30
}

Remove-Item "eas_downloader.py" -Force -ErrorAction SilentlyContinue
Write-Host "`n[DONE]" -ForegroundColor Green