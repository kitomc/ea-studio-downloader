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

$script:exitCode = 0
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

function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [int]$MaxRetries = 2,
        [int]$DelaySeconds = 3,
        [string]$Label = "operacion"
    )
    for ($i = 0; $i -le $MaxRetries; $i++) {
        try {
            return & $Action
        } catch {
            if ($i -lt $MaxRetries) {
                Write-Warn "${Label}: intento $($i+1)/$($MaxRetries+1) fallo. Reintentando en ${DelaySeconds}s..."
                Write-Warn "  Error: $($_.Exception.Message)"
                Start-Sleep -Seconds $DelaySeconds
            } else {
                Write-Err "${Label}: todos los intentos fallaron tras $($MaxRetries+1) intentos"
                throw
            }
        }
    }
}

function Write-Log($path, $msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $path -Value "[$ts] $msg"
}

function Write-ErrorLog($path, $context, $msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $path -Value "[$ts] [$context] $msg"
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
    $ghAvailable = $null
    try { $ghAvailable = Get-Command gh -ErrorAction Stop } catch { $ghAvailable = $null }

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
        $script:exitCode = 1; exit $script:exitCode
    }
}

if ([string]::IsNullOrEmpty($userDataDir)) {
    Write-Err "Error interno: variable de directorio de navegador no inicializada."
    $script:exitCode = 1; exit $script:exitCode
}

if (-not (Test-Path $userDataDir)) {
    Write-Err "No se encontro el directorio de datos de ${browserName} en: $userDataDir"
    Write-Err "Asegurate de tener ${browserName} instalado y haberlo abierto al menos una vez."
    $script:exitCode = 1; exit $script:exitCode
}

Write-Ok "Navegador: $browserName"
Write-Info "User Data: $userDataDir"

# Seleccion de perfil — detecta y pregunta al usuario
function Select-Profile {
    param(
        [array]$DetectedProfiles,
        [string]$BrowserName,
        [string]$DefaultProfile = "Default"
    )

    if (-not $DetectedProfiles -or $DetectedProfiles.Count -eq 0) {
        Write-Warn "No se pudieron detectar perfiles automaticamente."
        $manual = Read-Host "Ingresa el nombre del perfil (Enter para '${DefaultProfile}')"
        return $(if ([string]::IsNullOrEmpty($manual)) { $DefaultProfile } else { $manual })
    }

    Write-Host ""
    Write-Host "Perfiles detectados en ${BrowserName}:" -ForegroundColor White
    for ($i = 0; $i -lt $DetectedProfiles.Count; $i++) {
        $display = if ($DetectedProfiles[$i].Name) { $DetectedProfiles[$i].Name } else { $DetectedProfiles[$i] }
        $detail = if ($DetectedProfiles[$i].Directory) { " ($($DetectedProfiles[$i].Directory))" } else { "" }
        Write-Host "  [$i] ${display}${detail}" -ForegroundColor White
    }

    $choice = Read-Host "Elige el numero de perfil, o ingresa el nombre manualmente (Enter para '${DefaultProfile}')"

    if ([string]::IsNullOrEmpty($choice)) {
        return $DefaultProfile
    }

    # Si es un numero valido dentro del rango, usarlo como indice
    if ($choice -match "^\d+$") {
        $num = [int]$choice
        if ($num -ge 0 -and $num -lt $DetectedProfiles.Count) {
            if ($DetectedProfiles[$num].Directory) {
                return $DetectedProfiles[$num].Directory
            }
            return $DetectedProfiles[$num]
        }
    }

    # Si no, usar el texto ingresado como nombre de perfil
    return $choice
}

$profileDir = "Default"

if ($channel -eq "firefox") {
    $firefoxProfiles = Get-ChildItem $userDataDir -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "\.default" -or $_.Name -match "\.release" }
    if (-not $firefoxProfiles -or $firefoxProfiles.Count -eq 0) {
        $firefoxProfiles = Get-ChildItem $userDataDir -Directory -ErrorAction SilentlyContinue
    }
    if (-not $firefoxProfiles -or $firefoxProfiles.Count -eq 0) {
        Write-Err "No se encontraron perfiles de Firefox en: $userDataDir"
        $script:exitCode = 1; exit $script:exitCode
    } elseif ($firefoxProfiles.Count -eq 1) {
        $userDataDir = $firefoxProfiles[0].FullName
        $profileDir = $firefoxProfiles[0].Name
        Write-Ok "Perfil Firefox detectado: $profileDir"
        $manual = Read-Host "Usar perfil '${profileDir}'? (Enter=si, N para ingresar otro)"
        if ($manual -eq "N" -or $manual -eq "n") {
            $profileDir = Select-Profile -DetectedProfiles ($firefoxProfiles | ForEach-Object { $_.Name }) -BrowserName "Firefox" -DefaultProfile $profileDir
            $selectedProfile = $firefoxProfiles | Where-Object { $_.Name -eq $profileDir -or $_.Name -like "*$profileDir*" }
            if ($selectedProfile) {
                $userDataDir = $selectedProfile.FullName
            }
        }
    } else {
        $profileChoice = Select-Profile -DetectedProfiles ($firefoxProfiles | ForEach-Object { $_ | Select-Object @{N="Name";E={$_.Name}} }) -BrowserName "Firefox"
        $selectedProfile = $firefoxProfiles | Where-Object { $_.Name -eq $profileChoice -or $_.Name -like "*$profileChoice*" }
        if ($selectedProfile) {
            $profileDir = $selectedProfile.Name
            $userDataDir = $selectedProfile.FullName
        } else {
            Write-Warn "Perfil '${profileChoice}' no encontrado exactamente. Usando como nombre literal."
            $profileDir = $profileChoice
        }
    }
} else {
    $profiles = Get-BrowserProfiles $userDataDir
    $profileDir = Select-Profile -DetectedProfiles $profiles -BrowserName $browserName

    # Verificar que el perfil existe en disco
    $profilePath = Join-Path $userDataDir $profileDir
    if (-not (Test-Path $profilePath)) {
        Write-Warn "El directorio de perfil '${profileDir}' no existe en ${userDataDir}"
        $fallback = Read-Host "Ingresa el nombre correcto del perfil (Enter para 'Default')"
        if (-not [string]::IsNullOrEmpty($fallback)) {
            $profileDir = $fallback
        } else {
            $profileDir = "Default"
        }
    }
}

Write-Ok "Perfil seleccionado: ${profileDir}"

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

    if ([string]::IsNullOrEmpty($githubUrl)) {
        Write-Err "No se ingreso ninguna URL."
        $script:exitCode = 1; exit $script:exitCode
    }

    $repoSlug = $null
    if ($githubUrl -match "github\.com/([^/]+/[^/]+?)(?:\.git|/|$)") {
        $repoSlug = $Matches[1]
    } elseif ($githubUrl -match "^[\w-]+/[\w.-]+$") {
        $repoSlug = $githubUrl
    }

    if ($githubUrl -match "raw\.githubusercontent\.com" -or ($githubUrl -match "\.json$" -and -not $repoSlug)) {
        $tempDir = Join-Path $env:TEMP "eas-launcher"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        $fileName = Split-Path $githubUrl -Leaf
        $filePath = Join-Path $tempDir $fileName

        Write-Info "Descargando $fileName desde URL directa..."
        try {
            Invoke-WebRequest -Uri $githubUrl -OutFile $filePath -UseBasicParsing
            if (-not (Test-Path $filePath) -or (Get-Item $filePath).Length -eq 0) {
                throw "Archivo descargado vacio o no existe"
            }
            $CollectionFile = $filePath
            Write-Ok "Descargado: $filePath"
        } catch {
            Write-Err "Error descargando desde GitHub: $_"
            $script:exitCode = 1; exit $script:exitCode
        }
    } elseif ($githubUrl -match "\.json$" -and $repoSlug) {
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

        Write-Info "Descargando $fileName desde $repoSlug..."
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
            if (-not (Test-Path $filePath) -or (Get-Item $filePath).Length -eq 0) {
                throw "El archivo descargado esta vacio o no existe"
            }
            $CollectionFile = $filePath
            Write-Ok "Descargado: $filePath"
        } catch {
            Write-Warn "gh api fallo, intentando con URL directa: $($_.Exception.Message)"
            try {
                $branch = "main"
                if ($githubUrl -match "github\.com/[^/]+/[^/]+/blob/([^/]+)/") {
                    $branch = $Matches[1]
                }
                $rawUrl = "https://raw.githubusercontent.com/$repoSlug/$branch/$jsonPath"
                Invoke-WebRequest -Uri $rawUrl -OutFile $filePath -UseBasicParsing
                $CollectionFile = $filePath
                Write-Ok "Descargado via raw URL: $filePath"
            } catch {
                Write-Err "No se pudo descargar el archivo: $_"
                $script:exitCode = 1; exit $script:exitCode
            }
        }
    } elseif ($repoSlug) {
        $repoName = Split-Path $repoSlug -Leaf
        $cloneDir = Join-Path $env:TEMP "eas-launcher\$repoName"

        if (Test-Path $cloneDir) {
            Write-Info "Actualizando repo existente..."
            Push-Location $cloneDir
            try {
                if ($ghAvailable) { gh repo sync 2>&1 | Out-Null }
                git pull --ff-only 2>&1 | Out-Null
            } catch {
                git pull 2>&1 | Out-Null
            }
            Pop-Location
        } else {
            Write-Info "Clonando repo..."
            New-Item -ItemType Directory -Path (Split-Path $cloneDir -Parent) -Force | Out-Null
            try {
                if ($ghAvailable) {
                    gh repo clone $repoSlug $cloneDir 2>&1
                } else {
                    throw "gh CLI no disponible"
                }
            } catch {
                Write-Warn "gh repo clone fallo, intentando con git clone..."
                try {
                    git clone "https://github.com/$repoSlug.git" $cloneDir 2>&1 | Out-Null
                } catch {
                    Write-Err "No se pudo clonar el repo: $_"
                    $script:exitCode = 1; exit $script:exitCode
                }
            }
            Write-Ok "Repo clonado: $repoSlug"
        }

        $jsonFiles = Get-ChildItem $cloneDir -Filter "*.json" -Recurse -Depth 2 -ErrorAction SilentlyContinue
        if (-not $jsonFiles -or $jsonFiles.Count -eq 0) {
            Write-Err "No se encontraron archivos .json en el repo clonado"
            $script:exitCode = 1; exit $script:exitCode
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
            if ([string]::IsNullOrEmpty($fileChoice) -or [int]$fileChoice -lt 0 -or [int]$fileChoice -ge $jsonFiles.Count) {
                Write-Err "Seleccion invalida"
                $script:exitCode = 1; exit $script:exitCode
            }
            $CollectionFile = $jsonFiles[[int]$fileChoice].FullName
            Write-Ok "Seleccionado: $($jsonFiles[[int]$fileChoice].Name)"
        }
    } else {
        Write-Err "No se pudo interpretar la URL. Usa formato: owner/repo o URL completa de GitHub"
        $script:exitCode = 1; exit $script:exitCode
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
        $script:exitCode = 1; exit $script:exitCode
    }
    $CollectionFile = $openFileDialog.FileName
}

if ([string]::IsNullOrEmpty($CollectionFile) -or -not (Test-Path $CollectionFile)) {
    Write-Err "El archivo de collection no existe o no es accesible: $CollectionFile"
    $script:exitCode = 1; exit $script:exitCode
}

Write-Ok "Collection: $CollectionFile"

# =============================================================================
# 5. VERIFICAR PREREQUISITOS (Fase 1)
# =============================================================================

Write-Step "Fase 1/6: Verificando prerequisitos"

try {
    Invoke-WithRetry -Action {
        $ver = python --version 2>&1
        if ($LASTEXITCODE -ne 0 -or -not $ver) { throw "python --version fallo" }
        Write-Ok "Python disponible: $($ver.Trim())"
    } -Label "Verificar Python" -MaxRetries 1 -DelaySeconds 2
} catch {
    Write-Err "Python no esta instalado. Instalalo desde python.org y asegurate de que este en el PATH."
    Remove-Item $pyScriptPath -Force -ErrorAction SilentlyContinue
    $script:exitCode = 1; exit $script:exitCode
}

try {
    Invoke-WithRetry -Action {
        python -c "import playwright" 2>$null
        if ($LASTEXITCODE -ne 0) {
            Write-Info "Instalando playwright..."
            pip install playwright -q 2>$null
            if ($LASTEXITCODE -ne 0) { throw "pip install playwright fallo" }
            playwright install chromium 2>$null
            if ($LASTEXITCODE -ne 0) { throw "playwright install chromium fallo" }
        }
    } -Label "Verificar/instalar Playwright" -MaxRetries 1 -DelaySeconds 5
} catch {
    Write-Err "No se pudo instalar Playwright: $_"
    Write-Warn "Intenta manualmente: pip install playwright && playwright install chromium"
    $script:exitCode = 1; exit $script:exitCode
}

Write-Ok "Playwright disponible"

# =============================================================================
# 6. DIRECTORIO DE DESCARGAS
# =============================================================================

$downloadBase = Join-Path $scriptDir "downloads"
New-Item -ItemType Directory -Path $downloadBase -Force | Out-Null

# =============================================================================
# 7. LEER COLLECTION (Fase 2)
# =============================================================================

Write-Step "Fase 2/6: Leyendo collection"

try {
    $collectionJson = Get-Content $CollectionFile -Raw -Encoding UTF8
    if ([string]::IsNullOrEmpty($collectionJson)) {
        throw "El archivo JSON esta vacio"
    }
    $strategyCount = ($collectionJson | ConvertFrom-Json).Count
    if ($strategyCount -eq 0) {
        throw "La collection no contiene estrategias (array vacio)"
    }
} catch {
    Write-Err "Error leyendo el archivo de collection: $_"
    Write-Err "Asegurate de que el archivo sea un JSON valido con un array de estrategias."
    $script:exitCode = 1; exit $script:exitCode
}

Write-Info "Collection tiene $strategyCount estrategias"
Write-Info "Descargas en: $downloadBase"

# Logging dual: log.txt + errors.log
$logFile = Join-Path $downloadBase "log.txt"
$errorsLogFile = Join-Path $downloadBase "errors.log"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Inicializar logs
"[$timestamp] === INICIO DESCARGA ===" | Out-File -FilePath $logFile -Encoding UTF8
"[$timestamp] === INICIO DESCARGA ===" | Out-File -FilePath $errorsLogFile -Encoding UTF8

Write-Log $logFile "Navegador: ${browserName} ($channel)"
Write-Log $logFile "Perfil: $profileDir"
Write-Log $logFile "Collection: $CollectionFile"
Write-Log $logFile "Estrategias esperadas: $strategyCount"
Write-Log $logFile "Directorios: $downloadBase"

# =============================================================================
# 8. GENERAR Y EJECUTAR SCRIPT PYTHON (Fase 3-5)
# =============================================================================

Write-Step "Fase 3/6: Preparando navegador y sesion EAS"

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
    if ([string]::IsNullOrEmpty($braveExePath)) {
        Write-Err "No se encontro Brave.exe. Buscado en: $($braveExePaths -join ', ')"
        $script:exitCode = 1; exit $script:exitCode
    }
}

$pyUserDataDir = $userDataDir -replace "\\", "\\"
$pyCollectionFile = $CollectionFile -replace "\\", "\\"
$pyDownloadBase = $downloadBase -replace "\\", "\\"
$pyBraveExePath = $braveExePath -replace "\\", "\\"
$pyLogFile = $logFile -replace "\\", "\\"
$pyErrorsLogFile = $errorsLogFile -replace "\\", "\\"

$MAX_CONSECUTIVE = 3

$pyScript = @"
import asyncio, json, os, sys, time
from pathlib import Path
from playwright.async_api import async_playwright

EDGE_USER_DATA = r"$pyUserDataDir"
COLLECTION_FILE = r"$pyCollectionFile"
DOWNLOAD_DIR = Path(r"$pyDownloadBase")
LOG_FILE = Path(r"$pyLogFile")
ERRORS_LOG = Path(r"$pyErrorsLogFile")
EXPECTED_TOTAL = $strategyCount
BROWSER_TYPE = "$browserType"
CHANNEL = "$channel"
PROFILE_DIR = "$profileDir"
BRAVE_EXE = r"$pyBraveExePath"
MAX_CONSECUTIVE_FAILURES = $MAX_CONSECUTIVE

def log(msg):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def log_error(context, msg):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{context}] {msg}"
    print(f"  [ERROR] {msg}")
    with open(ERRORS_LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

async def with_retry(coro_factory, max_retries=2, delay=3, label=""):
    for i in range(max_retries + 1):
        try:
            return await coro_factory()
        except Exception as e:
            if i < max_retries:
                log(f"[RETRY] {label}: intento {i+1}/{max_retries+1} - {e}")
                log_error("RETRY", f"{label}: {e}")
                await asyncio.sleep(delay)
            else:
                log(f"[FAIL] {label}: todos los intentos fallaron tras {max_retries+1}")
                log_error("FAIL", f"{label}: {e}")
                raise

async def main():
    py_exit_code = 0
    log(f"Navegador: {BROWSER_TYPE} / channel={CHANNEL} / perfil={PROFILE_DIR}")

    import subprocess
    process_name = "${processName}"
    log(f"Cerrando procesos de {process_name}...")
    subprocess.run(["taskkill", "/F", "/IM", f"{process_name}.exe"], capture_output=True)
    await asyncio.sleep(2)

    async with async_playwright() as p:
        # --- FASES 3-4: Lanzar navegador ---
        try:
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
        except Exception as e:
            log_error("FATAL", f"Error lanzando navegador {CHANNEL}: {e}")
            return 1

        page = browser.pages[0] if browser.pages else await browser.new_page()

        # --- FASE 3: Navegar y verificar sesion ---
        try:
            await with_retry(
                lambda: page.goto("https://expert-advisor-studio.com/#", wait_until="networkidle", timeout=30000),
                max_retries=1, delay=3, label="Cargar EAS"
            )
            await page.wait_for_timeout(2000)
        except Exception as e:
            log_error("FATAL", f"No se pudo cargar expert-advisor-studio.com: {e}")
            await browser.close()
            return 1

        try:
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
                log_error("FATAL", "No hay sesion activa. Logueate en EAS y reintenta.")
                await browser.close()
                return 1
        except Exception as e:
            log_error("FATAL", f"Error verificando sesion: {e}")
            await browser.close()
            return 1

        log("Sesion verificada. Iniciando descarga...")

        # --- FASE 4: Upload Collection ---
        try:
            await with_retry(
                lambda: page.click("#eas-navbar-collection-link"),
                max_retries=1, delay=2, label="Click navbar collection"
            )
            await page.wait_for_load_state("networkidle", timeout=15000)
            await page.wait_for_timeout(1500)
        except Exception as e:
            log_error("FATAL", f"No se pudo navegar a Collection: {e}")
            await browser.close()
            return 1

        log("Navegando a Collection...")

        try:
            await with_retry(
                lambda: _upload_collection(page),
                max_retries=1, delay=3, label="Upload collection"
            )
        except Exception as e:
            log_error("FATAL", f"No se pudo subir la collection: {e}")
            await browser.close()
            return 1

        # --- FASE 5: Loop de descarga ---
        log(f"Descargando {EXPECTED_TOTAL} estrategias...")

        first_record = await page.query_selector('[id^="collection-record-"]')
        if not first_record:
            log_error("FATAL", "No hay registros en la collection.")
            await browser.close()
            return 1

        await first_record.click()
        try:
            await page.wait_for_selector("#editor-toolbar-export", timeout=10000)
            await page.wait_for_timeout(500)
        except:
            log_error("FATAL", "Editor no cargo.")
            await browser.close()
            return 1

        await page.wait_for_timeout(500)

        seen_ids = set()
        total_ok = 0
        total_mq4 = 0
        total_csv = 0
        total_err = 0
        errors_log = []
        consecutive_same = 0
        consecutive_failures = 0
        start_time = time.time()

        strategy_num = 0
        last_id = None

        while strategy_num < EXPECTED_TOTAL:
            strategy_num += 1
            strategy_errored = False
            mq5_stem = None

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
            log(f"[{strategy_num}/{EXPECTED_TOTAL}] {current_id or 'N/A'} (~{eta:.0f}s restantes)")

            # --- MQ4 ---
            try:
                await with_retry(
                    lambda: _download_mq4(page, DOWNLOAD_DIR, current_id or strategy_num),
                    max_retries=1, delay=2, label=f"MQ4 #{strategy_num}"
                )
                mq4_result = await _download_mq4(page, DOWNLOAD_DIR, current_id or strategy_num)
                mq4_name, mq5_stem = mq4_result
                log(f"  [OK] MQ4: {mq4_name}")
                total_mq4 += 1
            except Exception as e:
                log_error("MQ4", f"Estrategia {strategy_num}: {e}")
                total_err += 1
                errors_log.append(f"{strategy_num}: MQ4 - {e}")
                strategy_errored = True

            # --- MQ5 ---
            if mq5_stem:
                try:
                    await with_retry(
                        lambda: _download_mq5(page, DOWNLOAD_DIR, seen_ids, mq5_stem, strategy_num),
                        max_retries=1, delay=2, label=f"MQ5 #{strategy_num}"
                    )
                    mq5_result = await _download_mq5(page, DOWNLOAD_DIR, seen_ids, mq5_stem, strategy_num)
                    mq5_name, is_dup = mq5_result
                    if is_dup:
                        strategy_num -= 1
                    else:
                        log(f"  [OK] MQ5: {mq5_name}")
                        total_ok += 1
                except Exception as e:
                    log_error("MQ5", f"Estrategia {strategy_num}: {e}")
                    total_err += 1
                    errors_log.append(f"{strategy_num}: MQ5 - {e}")
                    strategy_errored = True

            # --- CSV ---
            if mq5_stem:
                try:
                    await with_retry(
                        lambda: _download_csv(page, DOWNLOAD_DIR, mq5_stem, strategy_num),
                        max_retries=1, delay=2, label=f"CSV #{strategy_num}"
                    )
                    csv_name = await _download_csv(page, DOWNLOAD_DIR, mq5_stem, strategy_num)
                    if csv_name:
                        log(f"  [OK] CSV: {csv_name}")
                        total_csv += 1
                except Exception as e:
                    log(f"  [--] CSV: {e}")

            # --- Volver al Editor ---
            try:
                editor_link = await page.query_selector("#eas-navbar-editor-link")
                if editor_link:
                    await editor_link.click()
                    await page.wait_for_timeout(500)
            except:
                pass

            # --- Control de fallos consecutivos ---
            if strategy_errored:
                consecutive_failures += 1
                log(f"  [WARN] {consecutive_failures}/{MAX_CONSECUTIVE_FAILURES} fallos consecutivos")
                if consecutive_failures >= MAX_CONSECUTIVE_FAILURES:
                    log(f"[STOP] {consecutive_failures} estrategias consecutivas fallaron. Deteniendo descarga.")
                    log_error("STOP", f"Detenido en estrategia {strategy_num} por {consecutive_failures} fallos consecutivos")
                    break
            else:
                consecutive_failures = 0

            # --- Ctrl+ArrowRight para siguiente ---
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
                    log_error("NAV", f"Estrategia {strategy_num}: Error navegando a siguiente: {e}")
                    break

        # --- Resumen ---
        elapsed = time.time() - start_time
        await browser.close()

        mq5_files = list(DOWNLOAD_DIR.glob("*.mq5"))
        mq4_files = list(DOWNLOAD_DIR.glob("*.mq4"))
        csv_files = list(DOWNLOAD_DIR.glob("*.csv"))

        completed_all = len(seen_ids) >= EXPECTED_TOTAL - 1  # -1 por la logica del break con ID repetido

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
                log_error("RESUMEN", err)
        log("=" * 55)

        # Codigo de salida
        if completed_all and total_err == 0:
            py_exit_code = 0
        elif completed_all and total_err > 0:
            py_exit_code = 3
        else:
            py_exit_code = 2

    return py_exit_code


async def _upload_collection(page):
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


async def _download_mq4(page, download_dir, strategy_id):
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
    await download.save_as(download_dir / mq4_name)
    return mq4_name, mq5_stem


async def _download_mq5(page, download_dir, seen_ids, mq5_stem, strategy_num):
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

    is_dup = mq5_stem_clean in seen_ids
    if is_dup:
        log(f"  [SKIP] Duplicado: {mq5_stem_clean}")
        await download.save_as(download_dir / mq5_name)
    else:
        seen_ids.add(mq5_stem_clean)
        await download.save_as(download_dir / mq5_name)

    return mq5_name, is_dup


async def _download_csv(page, download_dir, mq5_stem, strategy_num):
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
        await csv_download.save_as(download_dir / csv_name)
        return csv_name
    return None


sys.exit(asyncio.run(main()))
"@

$pyScriptPath = Join-Path $env:TEMP "eas_downloader_$(Get-Random).py"
$pyScript | Out-File -FilePath $pyScriptPath -Encoding UTF8

Write-Step "Fase 3b/6: Iniciando navegador con ${browserName}"
Write-Step "Fase 4/6: Subiendo collection"
Write-Step "Fase 5/6: Descargando $strategyCount estrategias"

$pyExitCode = 0
try {
    python $pyScriptPath
    $pyExitCode = $LASTEXITCODE
} catch {
    Write-Err "Error ejecutando script de descarga: $_"
    $pyExitCode = 1
}

# =============================================================================
# 9. VALIDACION (Fase 6)
# =============================================================================

Write-Step "Fase 6/6: Validacion de resultados"

$mq5Files = Get-ChildItem "$downloadBase\*.mq5" -ErrorAction SilentlyContinue
$mq4Files = Get-ChildItem "$downloadBase\*.mq4" -ErrorAction SilentlyContinue
$csvFiles  = Get-ChildItem "$downloadBase\*.csv"  -ErrorAction SilentlyContinue

$mq5Count = if ($mq5Files) { $mq5Files.Count } else { 0 }
$mq4Count = if ($mq4Files) { $mq4Files.Count } else { 0 }
$csvCount = if ($csvFiles) { $csvFiles.Count } else { 0 }

$pct = if ($strategyCount -gt 0) { [math]::Round($mq5Count / $strategyCount * 100, 1) } else { 0 }
$resultColor = if ($pct -ge 100) { "Green" } elseif ($pct -ge 50) { "Yellow" } else { "Red" }
Write-Host "[RESULTADO] MQ5: $mq5Count | MQ4: $mq4Count | CSV: $csvCount de $strategyCount ($pct%)" -ForegroundColor $resultColor

# Errores detectados
if (Test-Path $errorsLogFile) {
    $errCount = (Get-Content $errorsLogFile | Where-Object { $_ -match "\[FATAL\]|\[STOP\]|\[MQ4\]|\[MQ5\]|\[NAV\]" }).Count
    if ($errCount -gt 0) {
        Write-Warn "$errCount errores registrados en errors.log"
    }
}

if (Test-Path (Join-Path $downloadBase "log.txt")) {
    Write-Host ""
    Write-Host "[LOG] Ultimas 30 lineas de log.txt:" -ForegroundColor Cyan
    Get-Content (Join-Path $downloadBase "log.txt") -Tail 30
}

Remove-Item $pyScriptPath -Force -ErrorAction SilentlyContinue

# Determinar codigo de salida final
if ($pyExitCode -le 0 -and $pct -ge 100) {
    $script:exitCode = 0
    Write-Ok "Descarga completada al 100%"
} elseif ($pyExitCode -eq 2 -or ($pct -lt 100 -and $pct -gt 0)) {
    $script:exitCode = 2
    Write-Warn "Descarga detenida por fallos consecutivos o incompleta"
} elseif ($pyExitCode -eq 3) {
    $script:exitCode = 3
    Write-Warn "Descarga completa con errores no consecutivos"
} else {
    $script:exitCode = 1
}

Write-Host ""
Write-Ok "[DONE] (exit code: $script:exitCode)"
exit $script:exitCode