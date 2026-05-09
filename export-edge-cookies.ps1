# export-edge-cookies.ps1
# Exporta cookies + sessionStorage/localStorage de Edge usando el perfil del usuario.
# EAS no usa cookies tradicionales; la sesion se almacena en localStorage.
# Estrategia: usa Playwright con el perfil Edge del usuario directamente.

Write-Host "[INFO] Verificando prerequisitos..." -ForegroundColor Cyan

# Verificar Python
try { python --version | Out-Null } catch {
    Write-Host "[ERROR] Python no instalado" -ForegroundColor Red
    exit 1
}

# Verificar/instalar Playwright
python -c "import playwright" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[INFO] Instalando playwright..." -ForegroundColor Yellow
    pip install playwright -q 2>$null
    playwright install chromium 2>$null
}

# Detectar perfil Edge "Kito" (Default)
$edgeUserData = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
$localStatePath = "$edgeUserData\Local State"
$localState = Get-Content $localStatePath -Raw | ConvertFrom-Json

$kitoDir = $null
foreach ($prop in $localState.profile.info_cache.PSObject.Properties) {
    $name = $prop.Value.name
    if ($name -and $name -like "*Kito*") {
        $kitoDir = $prop.Name
        break
    }
}

if (-not $kitoDir) {
    $kitoDir = "Default"
    Write-Host "[INFO] Usando perfil Default (Kito)" -ForegroundColor Cyan
} else {
    Write-Host "[INFO] Perfil Kito encontrado: $kitoDir" -ForegroundColor Green
}

Write-Host "[INFO] Edge User Data: $edgeUserData" -ForegroundColor Cyan
Write-Host "[INFO] Perfil: $kitoDir" -ForegroundColor Cyan

# Generar script Python que usa Playwright con el perfil Edge del usuario (Kito = Default)
$pythonScript = @'
import asyncio, json, os, sys
from pathlib import Path

EDGE_USER_DATA = os.path.join(os.environ["LOCALAPPDATA"], "Microsoft", "Edge", "User Data")
PROFILE_NAME = "Default"  # Kito = Default
OUTPUT = os.path.join(os.getcwd(), "cookies.json")

async def main():
    from playwright.async_api import async_playwright
    
    async with async_playwright() as p:
        # Usar Edge con perfil Kito (Default)
        # user_data_dir DEBE ser el directorio padre de Edge, NO el subdirectorio Default
        # El perfil se selecciona via --profile-directory=Default
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
        
        # Verificar que estamos en el perfil Kito
        actual_profile = await page.evaluate("() => navigator.userAgent")
        print(f"[INFO] User Agent: {actual_profile[:80]}...")
        
        print("[INFO] Navegando a expert-advisor-studio.com...")
        await page.goto("https://expert-advisor-studio.com/", wait_until="networkidle", timeout=30000)
        await page.wait_for_timeout(3000)
        
        current_url = page.url
        page_title = await page.title()
        print(f"[INFO] URL actual: {current_url}")
        print(f"[INFO] Titulo: {page_title}")
        
        # Extraer localStorage (donde EAS guarda la sesion)
        local_storage_data = await page.evaluate("""() => {
            const data = {};
            for (let i = 0; i < localStorage.length; i++) {
                const key = localStorage.key(i);
                data[key] = localStorage.getItem(key);
            }
            return data;
        }""")
        
        # Extraer cookies del navegador
        cookies = await browser.cookies(["https://expert-advisor-studio.com", "https://www.expert-advisor-studio.com"])
        
        print(f"[INFO] Cookies del navegador para EAS: {len(cookies)}")
        print(f"[INFO] localStorage keys: {len(local_storage_data)}")
        print(f"[INFO] localStorage keys encontradas: {list(local_storage_data.keys())}")
        
        # Verificar sesion activa buscando indicadores en localStorage
        app_status = local_storage_data.get("eaStudio-app-status", "{}")
        if "Premium" in app_status or "premium" in app_status:
            print(f"[OK] Sesion PREMIUM detectada!")
        elif app_status != "{}":
            print(f"[INFO] Estado de app: {app_status[:100]}")
        
        playwright_cookies = []
        for c in cookies:
            playwright_cookies.append({
                "name": c.get("name", ""),
                "value": c.get("value", ""),
                "domain": c.get("domain", ""),
                "path": c.get("path", "/"),
                "expires": c.get("expires", -1),
                "httpOnly": c.get("httpOnly", False),
                "secure": c.get("secure", False),
                "sameSite": c.get("sameSite", "None"),
            })
        
        export_data = {
            "cookies": playwright_cookies,
            "localStorage": local_storage_data,
            "source_profile": PROFILE_NAME,
            "url": current_url,
            "title": page_title,
        }
        
        with open(OUTPUT, "w", encoding="utf-8") as f:
            json.dump(export_data, f, indent=2, ensure_ascii=False)
        
        print(f"[OK] Exportados: {len(playwright_cookies)} cookies, {len(local_storage_data)} localStorage keys")
        print(f"[OK] Datos guardados en {OUTPUT}")
        
        # Cerrar navegador - el downloader abrira una nueva sesion
        await browser.close()
        print("[DONE] Sesion exportada. Ejecuta ahora: .\\eas-downloader.ps1")

asyncio.run(main())
'@

$pythonScript | Out-File -FilePath "export_eas_session.py" -Encoding UTF8
Write-Host "[INFO] Ejecutando exportacion de sesion..." -ForegroundColor Cyan

python export_eas_session.py

if (Test-Path "cookies.json") {
    $data = Get-Content "cookies.json" | ConvertFrom-Json
    $cookieCount = $data.cookies.Count
    $lsCount = $data.localStorage.PSObject.Properties.Count
    Write-Host "[OK] cookies.json creado con $cookieCount cookies y $lsCount localStorage keys" -ForegroundColor Green
    Write-Host "[OK] URL verificada: $($data.url)" -ForegroundColor Green
    Write-Host "[OK] Titulo: $($data.title)" -ForegroundColor Green
    
    if ($cookieCount -eq 0 -and $lsCount -eq 0) {
        Write-Host "[WARN] No se encontraron cookies ni localStorage para EAS." -ForegroundColor Yellow
        Write-Host "[WARN] Verifica que estes logueado en expert-advisor-studio.com en Edge." -ForegroundColor Yellow
    }
} else {
    Write-Host "[ERROR] No se pudo crear cookies.json" -ForegroundColor Red
    exit 1
}

Remove-Item "export_eas_session.py" -Force -ErrorAction SilentlyContinue
Write-Host "[DONE] Sesion exportada. Ahora ejecuta: .\eas-downloader.ps1" -ForegroundColor Green