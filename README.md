# MOSCA Reporter Automation

Automatización de reportes de performance para Mosca Hnos. (Meta, Google Ads, TikTok y GA4), con generación de Excel, subida a Drive y envío de emails.

## Estructura del proyecto

- `reporter/reporte.py`: script principal de extracción, consolidación y envío.
- `reporter/.env.example`: plantilla de variables de entorno.
- `reporter/.env`: configuración local (credenciales y parámetros).
- `reporter/requirements.txt`: dependencias Python.
- `reporter/Guia_Credenciales.md`: guía detallada para obtener credenciales.
- `reporter/reporte.log`: log de ejecuciones automáticas.

## 1) Configuración inicial

### 1.1 Requisitos

- macOS con `python3` disponible.
- Accesos a Meta Ads, Google Ads, TikTok Ads, GA4, Gmail/Workspace.

### 1.2 Instalar dependencias

```bash
python3 -m pip install -r reporter/requirements.txt
```

### 1.3 Crear `.env`

```bash
cp reporter/.env.example reporter/.env
```

Completar variables en `reporter/.env` (ver detalle en `reporter/Guia_Credenciales.md`).

Variables clave:

- Meta: `META_ACCESS_TOKEN`, `META_AD_ACCOUNT_ID`
- Google Ads: `GOOGLE_DEVELOPER_TOKEN`, `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `GOOGLE_REFRESH_TOKEN`, `GOOGLE_CUSTOMER_ID`, `GOOGLE_MCC_ID`
- TikTok: `TIKTOK_ACCESS_TOKEN`, `TIKTOK_ADVERTISER_ID`
- GA4: `GA4_PROPERTY_ID`, `GA4_KEY_FILE`
- Email: `EMAIL_SENDER`, `EMAIL_APP_PASSWORD`, `EMAIL_RECIPIENT`
- IA (opcional): `ANTHROPIC_API_KEY`

### 1.4 Configurar GA4 con ruta absoluta

En `reporter/.env`, usar:

```env
GA4_KEY_FILE=/Users/cardinal/Documents/CARDINAL/MOSCA/reporter/ga4-key.json
```

## 2) Ejecución manual

### 2.1 Modo semanal (sin argumentos)

```bash
python3 reporter/reporte.py
```

- Usa período desde el día 1 del mes hasta **hoy**.
- Genera reporte semanal, status tab, subida a Drive y envíos de email.

### 2.2 Modo diario

```bash
python3 reporter/reporte.py --daily
```

- Usa período desde el día 1 del mes hasta **ayer**.
- Actualiza hojas diarias de Excel.

### 2.3 Modo manual con fechas

```bash
python3 reporter/reporte.py 2026-04-01 2026-04-17
```

## 3) Runner de cron (recomendado)

Para evitar errores de permisos y de contexto en cron, se usa wrapper con rutas absolutas:

- Script: `/Users/cardinal/.local/bin/mosca_report_runner.sh`

Contenido actual:

```zsh
#!/bin/zsh

PYTHON_BIN="/usr/bin/python3"
SCRIPT_PATH="/Users/cardinal/Documents/CARDINAL/MOSCA/reporter/reporte.py"
LOG_PATH="/Users/cardinal/Documents/CARDINAL/MOSCA/reporter/reporte.log"

if [[ "$#" -gt 0 ]]; then
  "$PYTHON_BIN" "$SCRIPT_PATH" "$@" >> "$LOG_PATH" 2>&1
else
  "$PYTHON_BIN" "$SCRIPT_PATH" >> "$LOG_PATH" 2>&1
fi
```

Permiso esperado:

```bash
chmod 755 /Users/cardinal/.local/bin/mosca_report_runner.sh
```

## 4) Programación activa (cron)

Entradas actuales:

```cron
45 9 * * * /Users/cardinal/.local/bin/mosca_report_runner.sh --daily
0 12 * * 1 /Users/cardinal/.local/bin/mosca_report_runner.sh
0 14 * * 5 /Users/cardinal/.local/bin/mosca_report_runner.sh
```

- Diario: 09:45 (modo `--daily`)
- Lunes: 12:00
- Viernes: 14:00

Verificar cron:

```bash
crontab -l
ps -ax -o pid,comm | awk '/cron/{print}'
```

## 5) Mantenimiento operativo

### 5.1 Revisar log

```bash
tail -n 120 reporter/reporte.log
```

### 5.2 Limpiar log

```bash
truncate -s 0 reporter/reporte.log
```

### 5.3 Probar exactamente como cron

```bash
/Users/cardinal/.local/bin/mosca_report_runner.sh --daily
```

### 5.4 Señales de ejecución correcta

En el log deben aparecer:

- `✓ Excel evolutivo guardado`
- `✓ Excel actualizado en Drive`
- `✓ Listo.`

## 6) Troubleshooting rápido

### Error: `Operation not permitted` en cron

Acciones:

1. Confirmar que cron use el runner (`mosca_report_runner.sh`) y no llame directo a `reporte.py`.
2. Verificar permiso ejecutable del runner (`chmod 755 ...`).
3. Verificar rutas absolutas del script y log en el runner.

### TikTok en 0 campañas

Si `TIKTOK_ACCESS_TOKEN` no está cargado o expiró, TikTok puede devolver 0 campañas y el resto del flujo sigue funcionando.

### Error de GA4 key file

Confirmar que `GA4_KEY_FILE` exista y sea absoluta en `reporter/.env`.

## 7) Comandos útiles de git

```bash
git status --short
git log -1 --oneline
git pull --rebase
git push
```
