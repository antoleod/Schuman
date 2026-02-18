# Configuracion

## `configs/appsettings.json`
- `ServiceNow`:
  - `BaseUrl`, `LoginUrl`, `WebViewProfileRoot`, timeouts y retries.
- `Excel`:
  - archivo/sheet por defecto y limites de escaneo.
- `Output`:
  - raiz `system`, subcarpetas `runs/logs/db`.
- `Documents`:
  - plantilla y carpeta de salida por defecto.

## Recomendaciones
- No hardcodear rutas en scripts; usar config.
- Mantener `Schuman List.xlsx` como fuente unica.
- Mantener `Reception_ITequipment.docx` en raiz o actualizar config.
