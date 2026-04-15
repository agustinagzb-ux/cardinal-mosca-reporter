# Guía de Credenciales — Reporte Automático
**Completar una vez, funciona para siempre.**

Hay 5 bloques: Meta, Google Ads, TikTok, GA4 y Email.
Al final de esta guía: cómo crear el archivo `.env` y activar el scheduler.

---

## PASO 0 — Crear el archivo .env

En la carpeta `reporter/`, copiar el archivo `.env.example` y renombrarlo `.env`:
```
cp .env.example .env
```
Luego ir completando cada sección con los valores de abajo.

---

## BLOQUE 1 — Meta Ads

### Qué necesitás
- `META_ACCESS_TOKEN` → token de acceso de larga duración
- `META_AD_ACCOUNT_ID` → ID de la cuenta publicitaria (formato `act_XXXXXXXXXX`)

### Cómo obtenerlos

**1. Encontrar el Ad Account ID:**
1. Ir a [business.facebook.com](https://business.facebook.com)
2. Configuración del negocio → Cuentas publicitarias
3. Hacer clic en la cuenta de Mosca Hnos.
4. El ID aparece como un número. Agregale `act_` adelante.
   → Ejemplo: si el ID es `123456789`, el valor es `act_123456789`

**2. Obtener el Access Token:**
1. Ir a [developers.facebook.com/tools/explorer](https://developers.facebook.com/tools/explorer)
2. Seleccionar tu app (o crear una nueva de tipo "Negocio")
3. En "User or Page" elegir tu usuario
4. Hacer clic en "Generate Access Token"
5. Tildar el permiso `ads_read`
6. Copiar el token generado

> ⚠️ Este token dura ~60 días. Para un token permanente, convertirlo en "long-lived token":
> Ir a: `https://graph.facebook.com/v19.0/oauth/access_token?grant_type=fb_exchange_token&client_id=TU_APP_ID&client_secret=TU_APP_SECRET&fb_exchange_token=TU_TOKEN_CORTO`

```
META_ACCESS_TOKEN=EAAxxxxxxxxxxxxxx...
META_AD_ACCOUNT_ID=act_123456789
```

---

## BLOQUE 2 — Google Ads

### Qué necesitás
- `GOOGLE_DEVELOPER_TOKEN` → token de desarrollador de Google Ads
- `GOOGLE_CLIENT_ID` y `GOOGLE_CLIENT_SECRET` → credenciales OAuth2
- `GOOGLE_REFRESH_TOKEN` → token de refresco permanente
- `GOOGLE_CUSTOMER_ID` → ID de la cuenta Google Ads de Mosca

### Cómo obtenerlos

**1. Developer Token:**
1. Iniciar sesión en [ads.google.com](https://ads.google.com) con la cuenta MCC (manager) o la cuenta de Mosca
2. Herramientas → Centro de API
3. Si no tenés developer token, solicitarlo. Para pruebas, el token en "test" alcanza.
4. Copiar el Developer Token

**2. Customer ID:**
1. En Google Ads, el ID de cliente aparece arriba a la derecha (formato `XXX-XXX-XXXX`)
2. Copiarlo sin guiones: `XXXXXXXXXX`

**3. OAuth2 (Client ID, Secret y Refresh Token):**
1. Ir a [console.cloud.google.com](https://console.cloud.google.com)
2. Crear un proyecto nuevo (o usar uno existente)
3. APIs y servicios → Habilitar APIs → buscar y habilitar **"Google Ads API"**
4. APIs y servicios → Credenciales → Crear credencial → **"ID de cliente de OAuth 2.0"**
5. Tipo de aplicación: **"Aplicación de escritorio"**
6. Descargar el JSON → abrirlo y copiar `client_id` y `client_secret`

**4. Obtener el Refresh Token (una sola vez):**

Ejecutar este comando en la terminal (reemplazar los valores):
```bash
python3 -c "
import requests, webbrowser, urllib.parse

CLIENT_ID = 'TU_CLIENT_ID'
CLIENT_SECRET = 'TU_CLIENT_SECRET'
SCOPE = 'https://www.googleapis.com/auth/adwords'
REDIRECT = 'urn:ietf:wg:oauth:2.0:oob'

auth_url = f'https://accounts.google.com/o/oauth2/auth?client_id={CLIENT_ID}&redirect_uri={REDIRECT}&scope={SCOPE}&response_type=code&access_type=offline'
print('Abrí esta URL en el navegador:')
print(auth_url)
"
```
1. Abrir la URL en el navegador, autorizar con la cuenta de Google Ads
2. Copiar el código que aparece
3. Ejecutar:
```bash
python3 -c "
import requests

r = requests.post('https://oauth2.googleapis.com/token', data={
    'code':          'PEGAR_EL_CODIGO',
    'client_id':     'TU_CLIENT_ID',
    'client_secret': 'TU_CLIENT_SECRET',
    'redirect_uri':  'urn:ietf:wg:oauth:2.0:oob',
    'grant_type':    'authorization_code',
})
print(r.json())
"
```
4. En la respuesta, copiar el valor de `refresh_token`

```
GOOGLE_DEVELOPER_TOKEN=xxxxxxxxxxxxxxxxxxxx
GOOGLE_CLIENT_ID=123456789-abc.apps.googleusercontent.com
GOOGLE_CLIENT_SECRET=GOCSPX-xxxxxxxxxx
GOOGLE_REFRESH_TOKEN=1//0xxxxxxxxxxxxxxxx
GOOGLE_CUSTOMER_ID=1234567890
```

---

## BLOQUE 3 — TikTok Ads

### Qué necesitás
- `TIKTOK_ACCESS_TOKEN`
- `TIKTOK_ADVERTISER_ID`

### Cómo obtenerlos

**1. Advertiser ID:**
1. Ir a [ads.tiktok.com](https://ads.tiktok.com)
2. En la esquina superior derecha, hacer clic en el nombre de la cuenta
3. El número que aparece es el Advertiser ID

**2. Access Token:**
1. Ir a [business-api.tiktok.com](https://business-api.tiktok.com) → My Apps
2. Si no tenés una app creada: Create App → tipo "Web" → llenar los campos mínimos
3. En la app creada: ir a **"Auth"** o **"Authorization"**
4. Hacer clic en **"Get Access Token"** (modo sandbox o producción)
5. Autorizar con la cuenta de TikTok Business
6. Copiar el `access_token` generado

> ⚠️ El token de TikTok dura 24 horas en modo sandbox y hasta 365 días en producción.
> Para producción, solicitar acceso de API estándar en la misma plataforma.

```
TIKTOK_ACCESS_TOKEN=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
TIKTOK_ADVERTISER_ID=1234567890123456789
```

---

## BLOQUE 4 — Google Analytics 4

### Qué necesitás
- `GA4_PROPERTY_ID` → número del Property de GA4
- `ga4-key.json` → archivo de Service Account

### Cómo obtenerlos

**1. Property ID de GA4:**
1. Ir a [analytics.google.com](https://analytics.google.com)
2. Administrar (ícono engranaje abajo a la izquierda)
3. En la columna "Propiedad" → Configuración de la propiedad
4. El **"ID de la propiedad"** es un número de 9 dígitos → copiarlo

**2. Crear Service Account y descargar el JSON:**
1. Ir a [console.cloud.google.com](https://console.cloud.google.com) → mismo proyecto que usaste en Google Ads
2. APIs y servicios → Habilitar **"Google Analytics Data API"**
3. IAM y administración → Cuentas de servicio → Crear cuenta de servicio
4. Nombre: `reporte-cardinal` → crear
5. En la cuenta creada → pestaña "Claves" → Agregar clave → JSON
6. Se descarga un archivo `xxxx.json` → renombrarlo `ga4-key.json` y copiarlo a la carpeta `reporter/`

**3. Dar acceso a la Service Account en GA4:**
1. En Google Analytics → Administrar → Acceso a la propiedad
2. Agregar usuario → pegar el email de la service account (termina en `@...iam.gserviceaccount.com`)
3. Rol: **"Lector"** → guardar

```
GA4_PROPERTY_ID=123456789
GA4_KEY_FILE=ga4-key.json
```

---

## BLOQUE 5 — Email (Gmail / Google Workspace)

### Qué necesitás
- `EMAIL_APP_PASSWORD` → contraseña de aplicación (NO la contraseña normal)

### Cómo obtenerla
1. Ir a [myaccount.google.com/security](https://myaccount.google.com/security)
2. Asegurarse de tener la **Verificación en 2 pasos activada**
3. Buscar **"Contraseñas de aplicaciones"**
4. Seleccionar: App = "Correo", Dispositivo = "Mac"
5. Google genera una contraseña de 16 caracteres → copiarla (se muestra una sola vez)

```
EMAIL_SENDER=agustina@cardinal.com.uy
EMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
EMAIL_RECIPIENT=agustina@cardinal.com.uy
```

---

## PASO FINAL — Activar el scheduler (cron)

Una vez que el `.env` esté completo y el script funcione, activar el scheduler automático:

**1. Probar que el script funciona:**
```bash
cd /Users/cardinal/Documents/CARDINAL/reporter
python3 reporte.py
```

**2. Abrir el crontab:**
```bash
crontab -e
```

**3. Agregar estas dos líneas:**
```
0 12 * * 1 /usr/bin/python3 /Users/cardinal/Documents/CARDINAL/reporter/reporte.py >> /Users/cardinal/Documents/CARDINAL/reporter/log.txt 2>&1
0 14 * * 5 /usr/bin/python3 /Users/cardinal/Documents/CARDINAL/reporter/reporte.py >> /Users/cardinal/Documents/CARDINAL/reporter/log.txt 2>&1
```

Guardar y cerrar (en vim: Escape → `:wq` → Enter).

**Significado:**
- `0 12 * * 1` → a las 12:00 todos los lunes
- `0 14 * * 5` → a las 14:00 todos los viernes
- El output se guarda en `log.txt` para revisar si algo falló

**4. Verificar que el cron quedó activo:**
```bash
crontab -l
```

> ⚠️ La Mac tiene que estar encendida y con internet al momento del horario programado.
> Si está apagada, el reporte no se envía ese día.

---

## Checklist de setup

- [ ] Crear `.env` desde `.env.example`
- [ ] Completar `META_ACCESS_TOKEN` y `META_AD_ACCOUNT_ID`
- [ ] Completar las 5 variables de Google Ads
- [ ] Completar `TIKTOK_ACCESS_TOKEN` y `TIKTOK_ADVERTISER_ID`
- [ ] Descargar `ga4-key.json` y copiarlo a la carpeta `reporter/`
- [ ] Dar acceso a la service account en GA4
- [ ] Completar `GA4_PROPERTY_ID`
- [ ] Generar App Password de Gmail y completar `EMAIL_APP_PASSWORD`
- [ ] Correr `python3 reporte.py` y confirmar que llega el mail
- [ ] Activar el cron
