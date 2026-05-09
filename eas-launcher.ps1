<# : CMD bootstrap — relanza en PowerShell si se ejecuta desde CMD
@echo off
powershell -ExecutionPolicy Bypass -NoProfile -File "%~f0" %*
exit /b
#>

# =============================================================================
# eas-launcher.ps1
# Descarga estrategias EAS (MQ4 + MQ5 + CSV) con soporte multi-navegador,
# collections desde GitHub (gh CLI), auto-actualizacion e instalacion como skill.
# =============================================================================

param(
    [switch]$Install,
    [switch]$SkipUpdate
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms

# =============================================================================
# FUNCIONES AUXILIARES
# =============================================================================

function Write-Step($msg) {
    Write-Host "`n[STEP] $msg" -ForegroundColor Cyan
}

function Write-Ok($msg) {
    Write-Host "[OK] $msg" -ForegroundColor Green
}

function Write-Warn($msg) {
    Write-Host "[WARN] $msg" -ForegroundColor Yellow
}

function Write-Err($msg) {
    Write-Host "[ERROR] $msg" -ForegroundColor Red
}

function Write-Info($msg) {
    Write-Host "[INFO] $msg" -ForegroundColor DarkGray
}

# =============================================================================
# 1. MODO INSTALAR — crear symlink en C:\Tools\
# =============================================================================

if ($Install) {
    Write-Step "Instalando eas-launcher como skill global..."

    $scriptPath = (Resolve-Path $PSCommandPath).Path
    $installDir = "C:\Tools"

    if (-not (Test-Path $installDir)) {
        New-Item -ItemType Directory -Path $installDir -Force | Out-Null
        Write-Ok "Creada carpeta $installDir"
    }

    $symlink = Join-Path $installDir "eas-launcher.ps1"

    if (Test-Path $symlink) {
        Remove-Item $symlink -Force
    }

    New-Item -ItemType SymbolicLink -Path $symlink -Target $scriptPath | Out-Null
    Write-Ok "Symlink creado: $symlink -> $scriptPath"

    $pathDirs = $env:PATH -split ";"
    if ($pathDirs -notcontains $installDir) {
        $oldPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
        if ($oldPath -notlike "*$installDir*") {
            try {
                [Environment]::SetEnvironmentVariable("PATH", "$oldPath;$installDir", "Machine")
                Write-Ok "$installDir agregado al PATH del sistema (requiere reiniciar terminal)"
            } catch {
                Write-Warn "No se pudo agregar al PATH del sistema (ejecuta como Admin)"
                Write-Info "Agrega manualmente: $installDir al PATH"
            }
        }
    } else {
        Write-Ok "$installDir ya esta en el PATH"
    }

    Write-Host ""
    Write-Ok "Instalacion completa. Ahora puedes ejecutar: eas-launcher"
    Write-Info "El symlink apunta al archivo real, asi que git pull o gh repo sync mantiene todo actualizado."
    exit 0
}

# =============================================================================
# 2. AUTO-ACTUALIZACION CON gh CLI
# =============================================================================

$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }

if (-not $SkipUpdate) {
    # Verificar si gh CLI esta disponible
    $ghAvailable = $null
    try { $ghAvailable = Get-Command gh -ErrorAction Stop } catch { $ghAvailable = $null }

    # Detectar si estamos en un repo git
    $inGitRepo = Test-Path (Join-Path $scriptDir ".git")

    if ($inGitRepo -and $ghAvailable) {
        Write-Step "Verificando actualizaciones con gh CLI..."
        Push-Location $scriptDir
        try {
            $ghResult = gh repo sync 2>&1
            if ($LASTEXITCODE -eq 0) {
                $updated = $false
                try {
                    $pullResult = git pull --ff-only 2>&1
                    if ($pullResult -notmatch "Already up to date") {
                        Write-Ok "Script actualizado desde GitHub: $pullResult"
                        $updated = $true
                    } else {
                        Write-Info "Script ya esta actualizado."
                    }
                } catch {
                    Write-Warn "git pull fallo: $_"
                }

                if ($updated) {
                    Write-Info "Re-ejecutando con la version nueva..."
                    Pop-Location
                    pwsh -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -SkipUpdate
                    exit 0
                }
            } else {
                Write-Warn "gh repo sync fallo, intentando git pull..."
                try {
                    $pullResult = git pull 2>&1
                    if ($pullResult -match "Already up to date") {
                        Write-Info "Script ya esta actualizado."
                    } else {
                        Write-Ok "Script actualizado via git pull."
                        Pop-Location
                        pwsh -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -SkipUpdate
                        exit 0
                    }
                } catch {
                    Write-Warn "No se pudo actualizar: $_"
                }
            }
        } catch {
            Write-Warn "gh no disponible o error: $_"
        }
        Pop-Location
    } elseif ($inGitRepo) {
        Write-Step "Verificando actualizaciones con git..."
        Push-Location $scriptDir
        try {
            $pullResult = git pull 2>&1
            if ($pullResult -match "Already up to date") {
                Write-Info "Script ya esta actualizado."
            } elseif ($pullResult -match "error") {
                Write-Warn "No se pudo actualizar: $pullResult"
            } else {
                Write-Ok "Script actualizado desde GitHub."
                Pop-Location
                pwsh -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -SkipUpdate
                exit 0
            }
        } catch {
            Write-Warn "git no disponible: $_"
        }
        Pop-Location
    }
}

# =============================================================================
# 3. DETECCION DE NAVEGADOR Y PERFIL
# =============================================================================

function Get-BrowserProfiles($userDataDir) {
    $localStatePath = Join-Path $userDataDir "Local State"
    if (-not (Test-Path $localStatePath)) {
        return @()
    }

    $profiles = @()

    try {
        $localState = Get-Content $localStatePath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($prop in $localState.profile.info_cache.PSObject.Properties) {
            $profileDir = $prop.Name
            $profileName = if ($prop.Value.name) { $prop.Value.name } else { $profileDir }
            $profiles += [PSCustomObject]@{
                Directory = $profileDir
                Name      = $profileName
            }
        }
    } catch {
        Write-Warn "No se pudo leer Local State de $userDataDir"
        if (Test-Path (Join-Path $userDataDir "Default")) {
            $profiles += [PSCustomObject]@{ Directory = "Default"; Name = "Default" }
        }
    }

    return $profiles
}

Write-Step "Seleccion de navegador"

Write-Host "En que navegador estas logueado en Expert Advisor Studio?" -ForegroundColor White
Write-Host "  [1] Microsoft Edge" -ForegroundColor White
Write-Host "  [2] Google Chrome" -ForegroundColor White
Write-Host "  [3] Firefox" -ForegroundColor White
Write-Host "  [4] Brave" -ForegroundColor White

$browserChoice = Read-Host "Elige (1-4)"

switch ($browserChoice) {
    "1" {
        $browserName = "Microsoft Edge"
        $userDataDir = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
        $channel = "msedge"
        $processName = "msedge"
    }
    "2" {
        $browserName = "Google Chrome"
        $userDataDir = "$env:LOCALAPPDATA\Google\Chrome\User Data"
        $channel = "chrome"
        $processName = "chrome"
    }
    "3" {
        $browserName = "Firefox"
        $userDataDir = "$env:APPDATA\Mozilla\Firefox\Profiles"
        $channel = "firefox"
        $processName = "firefox"
    }
    "4" {
        $browserName = "Brave"
        $userDataDir = "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data"
        $channel = "chrome"
        $processName = "brave"
    }
    default {
        Write-Err "Opcion invalida. Saliendo."
        exit 1
    }
}

if (-not (Test-Path $userDataDir)) {
    Write-Err "No se encontro el directorio de datos de ${browserName} en: $userDataDir"
    Write-Err "Asegurate de tener ${browserName} instalado y haberlo abierto al menos una vez."
    exit 1
}

Write-Ok "Navegador: $browserName"
Write-Info "User Data: $userDataDir"

# Seleccion de perfil
$profileDir = "Default"

if ($channel -eq "firefox") {
    $firefoxProfiles = Get-ChildItem $userDataDir -Directory | Where-Object { $_.Name -match "\.default" -or $_.Name -match "\.release" }
    if ($firefoxProfiles.Count -eq 0) {
        $firefoxProfiles = Get-ChildItem $userDataDir -Directory
    }
    if ($firefoxProfiles.Count -eq 1) {
        $userDataDir = $firefoxProfiles[0].FullName
        $profileDir = $firefoxProfiles[0].Name
        Write-Ok "Perfil Firefox: $profileDir"
    } elseif ($firefoxProfiles.Count -gt 1) {
        Write-Host ""
        Write-Host "Se encontraron multiples perfiles de Firefox:" -ForegroundColor White
        for ($i = 0; $i -lt $firefoxProfiles.Count; $i++) {
            Write-Host "  [$i] $($firefoxProfiles[$i].Name)" -ForegroundColor White
        }
        $profileChoice = Read-Host "Elige el perfil (0-$($firefoxProfiles.Count - 1))"
        $selected = $firefoxProfiles[[int]$profileChoice]
        $userDataDir = $selected.FullName
        $profileDir = $selected.Name
        Write-Ok "Perfil seleccionado: $profileDir"
    }
} else {
    $profiles = Get-BrowserProfiles $userDataDir

    if ($profiles.Count -eq 0) {
        Write-Info "Usando perfil Default"
        $profileDir = "Default"
    } elseif ($profiles.Count -eq 1) {
        $profileDir = $profiles[0].Directory
        Write-Ok "Perfil: $($profiles[0].Name) ($profileDir)"
    } else {
        Write-Host ""
        Write-Host "Se encontraron $($profiles.Count) perfiles en ${browserName}:" -ForegroundColor White
        for ($i = 0; $i -lt $profiles.Count; $i++) {
            Write-Host "  [$i] $($profiles[$i].Name) ($($profiles[$i].Directory))" -ForegroundColor White
        }
        $profileChoice = Read-Host "Elige el perfil (0-$($profiles.Count - 1))"
        if ([int]$profileChoice -ge 0 -and [int]$profileChoice -lt $profiles.Count) {
            $profileDir = $profiles[[int]$profileChoice].Directory
            Write-Ok "Perfil seleccionado: $($profiles[[int]$profileChoice].Name) ($profileDir)"
        } else {
            $profileDir = "Default"
            Write-Warn "Seleccion invalida, usando Default"
        }
    }
}

# =============================================================================
# 4. SELECCION DE COLLECTION (LOCAL O GITHUB via gh CLI)
# =============================================================================

Write-Step "Seleccion de Collection"

$CollectionFile = $null

$fromGithub = Read-Host "La collection esta en GitHub? (S/N)"

if ($fromGithub -eq "S" -or $fromGithub -eq "s") {
    $ghAvailable = $null
    try { $ghAvailable = Get-Command gh -ErrorAction Stop } catch { $ghAvailable = $null }

    $githubUrl = Read-Host "Ingresa la URL del repositorio o archivo JSON en GitHub (ej: owner/repo o URL completa)"

    # Parsear owner/repo desde diferentes formatos de URL
    $repoSlug = $null
    if ($githubUrl -match "github\.com/([^/]+/[^/]+?)(?:\.git|/|$)") {
        $repoSlug = $Matches[1]
    } elseif ($githubUrl -match "^[\w-]+/[\w.-]+$") {
        $repoSlug = $githubUrl
    }

    # Detectar si es una URL raw a un archivo JSON especifico
    if ($githubUrl -match "raw\.githubusercontent\.com" -or ($githubUrl -match "\.json$" -and -not $repoSlug)) {
        # Descargar directamente con Invoke-WebRequest
        $tempDir = Join-Path $env:TEMP "eas-launcher"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        $fileName = Split-Path $githubUrl -Leaf
        $filePath = Join-Path $tempDir $fileName

        Write-Info "Descargando $fileName desde URL directa..."
        try {
            Invoke-WebRequest -Uri $githubUrl -OutFile $filePath -UseBasicParsing
            $CollectionFile = $filePath
            Write-Ok "Descargado: $filePath"
        } catch {
            Write-Err "Error descargando desde GitHub: $_"
            exit 1
        }
    } elseif ($githubUrl -match "\.json$" -and $repoSlug) {
        # Es una URL a un archivo JSON dentro de un repo — usar gh api para obtener el contenido
        $jsonPath = ""
        if ($githubUrl -match "github\.com/[^/]+/[^/]+/blob/[^/]+/(.+\.json)") {
            $jsonPath = $Matches[1]
        } elseif ($githubUrl -match "/(.+\.json)$") {
            $jsonPath = $Matches[1]
        }

        $tempDir = Join-Path $env:TEMP "eas-launcher"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        $fileName = if ($jsonPath) { Split-Path $jsonPath -Leaf } else { "collection.json" }
        $filePath = Join-Path $tempDir $fileName

        Write-Info "Descargando $fileName desde $repoSlug con gh CLI..."
        try {
            if ($jsonPath) {
                gh api "/repos/$repoSlug/contents/$jsonPath" -q ".download_url" | ForEach-Object {
                    Invoke-WebRequest -Uri $_ -OutFile $filePath -UseBasicParsing
                }
            } else {
                $downloadUrl = gh api "/repos/$repoSlug/contents" -q ".[].download_url" 2>$null
                if ($downloadUrl) {
                    Invoke-WebRequest -Uri ($downloadUrl | Select-Object -First 1) -OutFile $filePath -UseBasicParsing
                }
            }
            if (Test-Path $filePath -and (Get-Item $filePath).Length -gt 0) {
                $CollectionFile = $filePath
                Write-Ok "Descargado: $filePath"
            } else {
                throw "El archivo descargado esta vacio o no existe"
            }
        } catch {
            Write-Warn "gh api fallo, intentando con URL directa..."
            try {
                $rawUrl = "https://raw.githubusercontent.com/$repoSlug/main/$jsonPath"
                Invoke-WebRequest -Uri $rawUrl -OutFile $filePath -UseBasicParsing
                $CollectionFile = $filePath
                Write-Ok "Descargado via raw URL: $filePath"
            } catch {
                Write-Err "No se pudo descargar el archivo: $_"
                exit 1
            }
        }
    } elseif ($repoSlug) {
        # Es un repo — clonar con gh repo clone
        $repoName = Split-Path $repoSlug -Leaf
        $cloneDir = Join-Path $env:TEMP "eas-launcher\$repoName"

        if (Test-Path $cloneDir) {
            Write-Info "Actualizando repo existente con gh..."
            Push-Location $cloneDir
            try {
                gh repo sync 2>&1 | Out-Null
                git pull --ff-only 2>&1 | Out-Null
            } catch {
                git pull 2>&1 | Out-Null
            }
            Pop-Location
        } else {
            Write-Info "Clonando repo con gh..."
            New-Item -ItemType Directory -Path (Split-Path $cloneDir -Parent) -Force | Out-Null
            try {
                gh repo clone $repoSlug $cloneDir 2>&1
                Write-Ok "Repo clonado: $repoSlug"
            } catch {
                Write-Warn "gh repo clone fallo, intentando con git clone..."
                try {
                    git clone "https://github.com/$repoSlug.git" $cloneDir 2>&1 | Out-Null
                } catch {
                    Write-Err "No se pudo clonar el repo: $_"
                    exit 1
                }
            }
        }

        # Buscar archivo JSON en el repo
        $jsonFiles = Get-ChildItem $cloneDir -Filter "*.json" -Recurse -Depth 2
        if ($jsonFiles.Count -eq 0) {
            Write-Err "No se encontraron archivos .json en el repo clonado"
            exit 1
        } elseif ($jsonFiles.Count -eq 1) {
            $CollectionFile = $jsonFiles[0].FullName
            Write-Ok "Collection: $($jsonFiles[0].Name)"
        } else {
            Write-Host ""
            Write-Host "Se encontraron $($jsonFiles.Count) archivos JSON:" -ForegroundColor White
            for ($i = 0; $i -lt $jsonFiles.Count; $i++) {
                Write-Host "  [$i] $($jsonFiles[$i].Name)" -ForegroundColor White
            }
            $fileChoice = Read-Host "Elige el archivo (0-$($jsonFiles.Count - 1))"
            if ([int]$fileChoice -ge 0 -and [int]$fileChoice -lt $jsonFiles.Count) {
                $CollectionFile = $jsonFiles[[int]$fileChoice].FullName
                Write-Ok "Seleccionado: $($jsonFiles[[int]$fileChoice].Name)"
            } else {
                Write-Err "Seleccion invalida"
                exit 1
            }
        }
    } else {
        Write-Err "No se pudo interpretar la URL. Usa formato: owner/repo o URL completa de GitHub"
        exit 1
    }
} else {
    Write-Info "Selecciona el archivo JSON de la collection..."
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "JSON Files|*.json|All Files|*.*"
    $openFileDialog.Title = "Selecciona la Collection de EAS"
    if (Test-Path $scriptDir) {
        $openFileDialog.InitialDirectory = $scriptDir
    }
    if ($openFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Err "No se selecciono ningun archivo."
        exit 1
    }
    $CollectionFile = $openFileDialog.FileName
}

Write-Ok "Collection: $CollectionFile"

# =============================================================================
# 5. VERIFICAR PREREQUISITOS
# =============================================================================

Write-Step "Verificando prerequisitos"

try { python --version | Out-Null } catch {
    Write-Err "Python no esta instalado. Instalalo desde python.org"
    exit 1
}
Write-Ok "Python disponible"

python -c "import playwright" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Info "Instalando playwright..."
    pip install playwright -q 2>$null
    playwright install chromium 2>$null
}
Write-Ok "Playwright disponible"

# =============================================================================
# 6. DIRECTORIO DE DESCARGAS
# =============================================================================

$downloadBase = Join-Path $scriptDir "downloads"
New-Item -ItemType Directory -Path $downloadBase -Force | Out-Null

# =============================================================================
# 7. LEER COLLECTION Y PREPARAR LOG
# =============================================================================

$collectionJson = Get-Content $CollectionFile -Raw
$strategyCount = ($collectionJson | ConvertFrom-Json).Count
Write-Info "Collection tiene $strategyCount estrategias"
Write-Info "Descargas en: $downloadBase"

$logFile = Join-Path $downloadBase "log.txt"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $logFile -Value "[$timestamp] === INICIO DESCARGA ==="
Add-Content -Path $logFile -Value "[$timestamp] Navegador: ${browserName} ($channel)"
Add-Content -Path $logFile -Value "[$timestamp] Perfil: $profileDir"
Add-Content -Path $logFile -Value "[$timestamp] Collection: $CollectionFile"
Add-Content -Path $logFile -Value "[$timestamp] Estrategias esperadas: $strategyCount"
Add-Content -Path $logFile -Value "[$timestamp] Directorio descargas: $downloadBase"

# =============================================================================
# 8. GENERAR Y EJECUTAR SCRIPT PYTHON
# =============================================================================

$browserType = if ($channel -eq "firefox") { "firefox" } else { "chromium" }

$braveExePath = ""
if ($browserChoice -eq "4") {
    $braveExePaths = @(
        "${env:ProgramFiles}\BraveSoftware\Brave-Browser\Application\brave.exe",
        "${env:ProgramFiles(x86)}\BraveSoftware\Brave-Browser\Application\brave.exe",
        "${env:LOCALAPPDATA}\BraveSoftware\Brave-Browser\Application\brave.exe"
    )
    foreach ($p in $braveExePaths) {
        if (Test-Path $p) {
            $braveExePath = $p
            break
        }
    }
}

$pyUserDataDir = $userDataDir -replace "\\", "\\"
$pyCollectionFile = $CollectionFile -replace "\\", "\\"
$pyDownloadBase = $downloadBase -replace "\\", "\\"
$pyBraveExePath = $braveExePath -replace "\\", "\\"

$pyScript = @"
import asyncio, json, os, sys, time
from pathlib import Path
from playwright.async_api import async_playwright

EDGE_USER_DATA = r"$pyUserDataDir"
COLLECTION_FILE = r"$pyCollectionFile"
DOWNLOAD_DIR = Path(r"$pyDownloadBase")
LOG_FILE = DOWNLOAD_DIR / "log.txt"
EXPECTED_TOTAL = $strategyCount
BROWSER_TYPE = "$browserType"
CHANNEL = "$channel"
PROFILE_DIR = "$profileDir"
BRAVE_EXE = r"$pyBraveExePath"

def log(msg):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

async def main():
    log(f"Navegador: {BROWSER_TYPE} / channel={CHANNEL} / perfil={PROFILE_DIR}")

    import subprocess
    process_name = "${processName}"
    log(f"Cerrando procesos de {process_name}...")
    subprocess.run(["taskkill", "/F", "/IM", f"{process_name}.exe"], capture_output=True)
    await asyncio.sleep(2)

    async with async_playwright() as p:
        launch_args = {
            "headless": False,
            "accept_downloads": True,
            "no_viewport": True,
        }

        if BROWSER_TYPE == "firefox":
            browser = await p.firefox.launch_persistent_context(
                user_data_dir=EDGE_USER_DATA,
                **launch_args,
            )
        elif CHANNEL == "chrome" and BRAVE_EXE:
            launch_args["executable_path"] = BRAVE_EXE
            launch_args["args"] = [
                f"--profile-directory={PROFILE_DIR}",
                "--disable-blink-features=AutomationControlled",
            ]
            browser = await p.chromium.launch_persistent_context(
                user_data_dir=EDGE_USER_DATA,
                **launch_args,
            )
        else:
            launch_args["channel"] = CHANNEL
            launch_args["args"] = [
                f"--profile-directory={PROFILE_DIR}",
                "--disable-blink-features=AutomationControlled",
            ]
            browser = await p.chromium.launch_persistent_context(
                user_data_dir=EDGE_USER_DATA,
                **launch_args,
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

        log("Navegando a Collection...")
        await page.click("#eas-navbar-collection-link")
        await page.wait_for_load_state("networkidle", timeout=15000)
        await page.wait_for_timeout(1500)

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

        await page.screenshot(path=str(DOWNLOAD_DIR / "debug_after_upload.png"))
        log("Screenshot guardado: debug_after_upload.png")

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

        log(f"Descargando {EXPECTED_TOTAL} estrategias...")

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

        await page.wait_for_timeout(500)

        strategy_num = 0
        last_id = None

        while strategy_num < EXPECTED_TOTAL:
            strategy_num += 1

            current_id = None
            try:
                id_el = await page.query_selector("#navbar-strategy-id")
                if id_el:
                    current_id = await id_el.inner_text()
            except:
                pass

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

            try:
                editor_link = await page.query_selector("#eas-navbar-editor-link")
                if editor_link:
                    await editor_link.click()
                    await page.wait_for_timeout(500)
            except:
                pass

            if strategy_num < EXPECTED_TOTAL:
                try:
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

$pyScriptPath = Join-Path $env:TEMP "eas_downloader_$(Get-Random).py"
$pyScript | Out-File -FilePath $pyScriptPath -Encoding UTF8

Write-Step "Iniciando descarga de $strategyCount estrategias con ${browserName}"
python $pyScriptPath

# =============================================================================
# 9. VALIDACION
# =============================================================================

Write-Step "Validacion de resultados"

$mq5Files = Get-ChildItem "$downloadBase\*.mq5" -ErrorAction SilentlyContinue
$mq4Files = Get-ChildItem "$downloadBase\*.mq4" -ErrorAction SilentlyContinue
$csvFiles  = Get-ChildItem "$downloadBase\*.csv"  -ErrorAction SilentlyContinue

$mq5Count = if ($mq5Files) { $mq5Files.Count } else { 0 }
$mq4Count = if ($mq4Files) { $mq4Files.Count } else { 0 }
$csvCount = if ($csvFiles) { $csvFiles.Count } else { 0 }

$pct = if ($strategyCount -gt 0) { [math]::Round($mq5Count / $strategyCount * 100, 1) } else { 0 }
Write-Host "[RESULTADO] MQ5: $mq5Count | MQ4: $mq4Count | CSV: $csvCount de $strategyCount ($pct%)" -ForegroundColor $(if ($pct -ge 100) {"Green"} else {"Yellow"})

if (Test-Path (Join-Path $downloadBase "log.txt")) {
    Write-Host ""
    Write-Host "[LOG] Ultimas 30 lineas de log.txt:" -ForegroundColor Cyan
    Get-Content (Join-Path $downloadBase "log.txt") -Tail 30
}

Remove-Item $pyScriptPath -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Ok "[DONE]"