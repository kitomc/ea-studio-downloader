# Error Handling — eas-launcher.ps1

## Resumen
Implementar manejo de errores estructurado en 6 fases con retry + stop, logging dual (`log.txt` + `errors.log`) y códigos de salida estandarizados.

## Fases

| Fase | Errores | Retry | Fatal |
|---|---|---|---|
| 1. Prerequisitos | Python, pip, Playwright | auto-install pip/playwright | sin Python → exit 1 |
| 2. Navegador | Perfil locked, no existe | taskkill + 3 intentos | 3 fallos → exit 1 |
| 3. Sesión EAS | Sin localStorage, page timeout | reload + 1 retry | sin sesión → exit 1 |
| 4. Upload Collection | File chooser, timeout | 2 intentos | si falla → exit 1 |
| 5. Descarga loop | MQ4/MQ5/CSV timeout, navigation | 2 retry/estrategia, stop tras 3 consecutivos | stop con resumen |
| 6. Validación | Siempre se ejecuta | — | nunca fatal |

## Códigos de salida
- `0`: 100% completado
- `1`: Error fatal (sin Python/navegador/sesión)
- `2`: Detenido por N fallos consecutivos
- `3`: Parcial con errores no consecutivos

## Logging
- `log.txt`: todo (timestamp + nivel + mensaje)
- `errors.log`: solo errores críticos con contexto (fase, estrategia, excepción)

## Función Invoke-WithRetry (PowerShell)
Wrapper genérico con `[scriptblock]$Action`, `$MaxRetries`, `$DelaySeconds`, `$Label`.

## Función with_retry (Python)
Wrapper asíncrono para operaciones en el loop de descarga.
