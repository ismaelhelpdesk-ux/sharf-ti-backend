# SHARF Devoluciones TI — Backend Python + PWA

## Arquitectura

```
┌─────────────────────────────────────────────────────────┐
│  CELULAR  →  PWA instalable (HTML + Service Worker)     │
│               ↕  HTTPS fetch() JSON                      │
│  SERVIDOR →  FastAPI Python                              │
│    ├── /api/sesion         Autenticar técnico            │
│    ├── /api/activos/{dni}  Buscar en Snipe-IT            │
│    ├── /api/qr             Decodificar QR (pyzbar)       │
│    ├── /api/devolucion     Proceso completo              │
│    │     ├── Snipe-IT checkin                            │
│    │     ├── PDF con ReportLab                           │
│    │     ├── Subir a Google Drive                        │
│    │     ├── Registrar en Sheets                         │
│    │     └── Enviar Gmail                                │
│    ├── /api/validar        Health-check                  │
│    └── /api/auditoria      Estadísticas                  │
└─────────────────────────────────────────────────────────┘
```

## Deploy en Railway (gratis hasta 500h/mes)

### 1. Preparar el repositorio
```bash
git init
git add .
git commit -m "SHARF API inicial"
```

### 2. Crear proyecto en Railway
1. Ve a [railway.app](https://railway.app) → New Project
2. Deploy from GitHub repo → selecciona tu repositorio
3. Railway detecta el Dockerfile automáticamente

### 3. Configurar variables de entorno en Railway
En el panel Variables, agrega:
```
SNIPE_TOKEN        = tu_token_de_snipe_it
GOOGLE_CREDS_JSON  = {"type":"service_account",...}  ← JSON completo
SHEET_LOG_ID       = 14LNid3_E7deg9rHD65TA9P_h8Lmr2s0HvWoejHlYQow
DRIVE_ACTAS_ROOT   = 0APknu_tBOg5SUk9PVA
GMAIL_USER         = helpdesk@holasharf.com
ENV                = production
SECRET_KEY         = un-secreto-largo-y-aleatorio
```

### 4. Obtener la URL
Railway asigna una URL automáticamente:
`https://sharf-ti-production.up.railway.app`

### 5. Instalar como app en el celular
1. Abre Chrome en Android / Safari en iPhone
2. Ve a `https://tu-url-railway.app`
3. Chrome: menú → "Agregar a pantalla de inicio"
4. Safari: compartir → "Añadir a pantalla de inicio"

¡La app aparece como ícono nativo sin App Store!

---

## Deploy en Google Cloud Run (gratis tier)

```bash
# Build y push
gcloud builds submit --tag gcr.io/TU_PROYECTO/sharf-ti

# Deploy
gcloud run deploy sharf-ti \
  --image gcr.io/TU_PROYECTO/sharf-ti \
  --platform managed \
  --region us-central1 \
  --allow-unauthenticated \
  --set-env-vars "ENV=production,SNIPE_TOKEN=..."
```

---

## Desarrollo local

```bash
# Instalar dependencias del sistema (Ubuntu/Debian)
sudo apt-get install libzbar0

# Instalar Python
pip install -r requirements.txt

# Copiar y editar variables
cp .env.example .env
# → edita .env con tus credenciales reales

# Arrancar el servidor
python main.py
# → http://localhost:8000
# → Docs: http://localhost:8000/docs
```

---

## Google Service Account (para Drive, Sheets, Gmail)

1. Ve a [console.cloud.google.com](https://console.cloud.google.com)
2. Crea un proyecto o usa el existente de GAS
3. Habilita las APIs:
   - Google Drive API
   - Google Sheets API
   - Gmail API
4. IAM → Service Accounts → Create
5. Descarga el JSON de credenciales
6. Comparte las carpetas Drive y el Sheet con el email de la SA
7. Pega el JSON completo en la variable `GOOGLE_CREDS_JSON`

---

## Estructura del proyecto

```
sharf_app/
├── main.py              ← FastAPI — endpoints principales
├── requirements.txt     ← Dependencias Python
├── Dockerfile           ← Deploy en cualquier servidor
├── railway.toml         ← Deploy específico Railway
├── .env.example         ← Variables de entorno (copiar a .env)
├── core/
│   ├── config.py        ← Configuración centralizada
│   ├── models.py        ← Modelos Pydantic (validación)
│   ├── snipe.py         ← Cliente Snipe-IT async
│   ├── sheets.py        ← Google Sheets
│   ├── drive.py         ← Google Drive
│   ├── mail.py          ← Gmail
│   ├── pdf.py           ← Generador PDF con ReportLab
│   ├── qr.py            ← Decodificador QR con pyzbar
│   └── auth.py          ← Lista de técnicos autorizados
└── static/
    ├── sw.js            ← Service Worker PWA
    ├── icon-192.png     ← Ícono app
    └── icon-512.png     ← Ícono app grande
```

---

## Ventajas sobre Google Apps Script

| Característica | GAS (antes) | Python FastAPI (ahora) |
|---|---|---|
| Tiempo de respuesta | 3-8 segundos | < 1 segundo |
| PDF | html2canvas (impreciso) | ReportLab (exacto) |
| QR | jsQR (JavaScript) | pyzbar (más preciso) |
| Offline | No | Sí (Service Worker) |
| Instalable como app | No | Sí (PWA) |
| Notificaciones push | No | Sí |
| Escalabilidad | Limitada (GAS cuotas) | Ilimitada |
| Debug | Logger.log | Logs estructurados |
| Testing | Manual | pytest automatizado |

---

## Endpoints disponibles

| Método | URL | Descripción |
|---|---|---|
| GET | `/` | Sirve la PWA |
| POST | `/api/sesion` | Autenticar técnico |
| GET | `/api/activos/{dni}` | Buscar activos Snipe-IT |
| GET | `/api/activo/{id}` | Detalle de activo |
| POST | `/api/qr` | Decodificar imagen QR |
| POST | `/api/devolucion` | Procesar devolución completa |
| GET | `/api/validar` | Health-check conexiones |
| GET | `/api/auditoria?meses=3` | Estadísticas de auditoría |
| GET | `/api/catalogos` | Áreas y sedes Snipe-IT |
| GET | `/manifest.json` | Manifiesto PWA |
| GET | `/docs` | Documentación interactiva Swagger |

---

## Contacto y soporte
Sistema desarrollado por Mesa de Servicio TI — Scharff
Administrador: ismael.helpdesk@holasharf.com
