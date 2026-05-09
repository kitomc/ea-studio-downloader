# EA Studio Downloader

Descarga automática de estrategias desde [Expert Advisor Studio](https://expert-advisor-studio.com/) en formato MQ4, MQ5 y CSV. Soporta colecciones de 300+ estrategias con navegación automatizada.

## Características

- **Multi-navegador**: Microsoft Edge, Google Chrome, Firefox, Brave
- **Detección automática de perfiles**: Lee los perfiles del navegador y permite seleccionar
- **Collections desde GitHub**: Descarga archivos `.json` directamente desde repositorios con `gh`
- **Auto-actualización**: Si el script está en un repo Git, ejecuta `gh repo sync` + `git pull` al inicio
- **Instalación global**: Symlink en `C:\Tools\` para ejecutar desde cualquier ruta
- **Compatible CMD y PowerShell**: Se auto-relanza en PowerShell si se ejecuta desde CMD
- **Descarga completa**: MQ4 + MQ5 + CSV + log.txt por cada estrategia

## Requisitos previos

| Requisito | Versión | Verificar |
|---|---|---|
| **Python** | 3.8+ | `python --version` |
| **PowerShell** | 5.1+ (Windows) o 7+ | `$PSVersionTable.PSVersion` |
| **Playwright** | Se instala automáticamente | `pip install playwright` |
| **GitHub CLI** (opcional) | Última estable | `gh --version` |
| **Git** (opcional) | 2.30+ | `git --version` |

### Instalación de dependencias

```powershell
# Python y Playwright (requerido)
pip install playwright
playwright install chromium

# GitHub CLI (opcional, para collections desde GitHub)
winget install GitHub.cli

# Git (opcional, para auto-actualización)
winget install Git.Git
```

## Instalación

### Opción 1: Clonar con Git

```powershell
git clone https://github.com/kitomc/ea-studio-downloader.git
cd ea-studio-downloader
```

### Opción 2: Descargar ZIP

1. Ir a [github.com/kitomc/ea-studio-downloader](https://github.com/kitomc/ea-studio-downloader)
2. Click en **Code** → **Download ZIP**
3. Extraer en la carpeta deseada

### Instalar como comando global

Esto crea un symlink en `C:\Tools\eas-launcher.ps1` y lo agrega al PATH del sistema. Requiere ejecutar PowerShell como **Administrador**:

```powershell
.\eas-launcher.ps1 -Install
```

Después de instalado, puedes ejecutar desde cualquier carpeta:

```powershell
eas-launcher
```

El symlink apunta al archivo original, así que `git pull` mantiene todo actualizado automáticamente.

## Uso

### Ejecución básica

```powershell
.\eas-launcher.ps1
```

O si está instalado globalmente:

```powershell
eas-launcher
```

### Desde CMD (Símbolo del sistema)

El script detecta si se ejecuta desde CMD y se relanza automáticamente en PowerShell:

```cmd
eas-launcher.ps1
```

### Parámetros

| Parámetro | Descripción |
|---|---|
| `-Install` | Instala el symlink global en `C:\Tools\` y lo agrega al PATH |
| `-SkipUpdate` | Omite la verificación de auto-actualización al inicio |

### Flujo interactivo

1. **Navegador**: Selecciona en qué navegador tienes sesión de EAS (Edge/Chrome/Firefox/Brave)
2. **Perfil**: Si hay múltiples perfiles, selecciona el correcto
3. **Collection**: Indica si el archivo `.json` está en GitHub o local
4. **Descarga**: El script abre el navegador, sube la collection y descarga MQ4+MQ5+CSV de cada estrategia

### Collection desde GitHub

CuandoSeleccionas "S" en la pregunta de GitHub, puedes ingresar:

- **URL de repo**: `kitomc/mis-collections` → clona con `gh repo clone`
- **URL de archivo raw**: `https://raw.githubusercontent.com/.../estrategias.json` → descarga directamente
- **URL completa de GitHub**: `https://github.com/kitomc/mis-collections/blob/main/estrategias.json` → extrae con `gh api`

### Auto-actualización

Si el script está en un repositorio Git (clonado con `git clone`), al iniciar ejecuta:

1. `gh repo sync` para sincronizar con GitHub
2. `git pull --ff-only` para traer los cambios locales

Si hay actualizaciones, se re-ejecuta automáticamente con la versión nueva.

Usa `-SkipUpdate` para omitir este paso:

```powershell
.\eas-launcher.ps1 -SkipUpdate
```

## Estructura de archivos

```
ea-studio-downloader/
├── eas-launcher.ps1        # Script principal unificado
├── eas-downloader.ps1      # Script original (solo Edge)
├── export-edge-cookies.ps1 # Exportador de sesión Edge
├── .gitignore
└── downloads/               # Carpeta de descargas (gitignored)
    ├── *.mq4
    ├── *.mq5
    ├── *.csv
    └── log.txt
```

## Cómo funciona

1. Abre el navegador seleccionado con tu perfil (sesión existente)
2. Navega a `expert-advisor-studio.com`
3. Verifica que hay sesión activa (localStorage `eaStudio-user`)
4. Va a Collection → Sube el archivo `.json`
5. Para cada estrategia:
   - Exporta **MQ4** → guarda en `downloads/`
   - Exporta **MQ5** → guarda en `downloads/`
   - Exporta **CSV** (Report → Journal → Export) → guarda en `downloads/`
   - Navega a la siguiente con `Ctrl+ArrowRight`
6. Genera `log.txt` con el resumen de la descarga

## Solución de problemas

### "No hay sesión activa"

Asegúrate de haber iniciado sesión en Expert Advisor Studio en el navegador seleccionado **antes** de ejecutar el script. El script usa tu perfil real del navegador.

### "Python no instalado"

Instala Python desde [python.org](https://www.python.org/downloads/) y asegúrate de marcar "Add to PATH" durante la instalación.

### "Playwright no encontrado"

```powershell
pip install playwright
playwright install chromium
```

### El navegador se cierra inmediatamente

Cierra todas las instancias del navegador antes de ejecutar el script. El script fuerza el cierre (`taskkill`) para evitar conflictos de perfil.

### Error de permisos al instalar con `-Install`

Ejecuta PowerShell como **Administrador**:

```powershell
Right-click PowerShell → "Ejecutar como administrador"
.\eas-launcher.ps1 -Install
```

### Firefox no funciona con Playwright

Playwright tiene soporte experimental para Firefox. Si falla:

```powershell
playwright install firefox
```

## Scripts incluidos

| Script | Descripción |
|---|---|
| `eas-launcher.ps1` | **Script principal**. Multi-navegador, gh CLI, auto-update, install global |
| `eas-downloader.ps1` | Script original. Solo Edge con perfil "Kito/Default" |
| `export-edge-cookies.ps1` | Exporta sesión de EAS desde localStorage de Edge |

## Licencia

Uso personal. Expert Advisor Studio es un producto de [EA Studio](https://expert-advisor-studio.com/).