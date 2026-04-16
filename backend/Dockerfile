# SHARF Devoluciones TI — Dockerfile
# Deploy: Railway, Google Cloud Run, Render, VPS con Docker

FROM python:3.12-slim

# Dependencias del sistema para pyzbar (lectura QR) y ReportLab
RUN apt-get update && apt-get install -y --no-install-recommends \
    libzbar0 \
    libzbar-dev \
    zlib1g-dev \
    libjpeg-dev \
    libfreetype6-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Instalar dependencias Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar código fuente
COPY . .

# Hacer ejecutable el script de arranque
RUN chmod +x start.sh

# Puerto de la aplicación
EXPOSE 8000

# Variables de entorno obligatorias (configurar en el servidor):
# SNIPE_TOKEN, GOOGLE_CREDS_JSON, SHEET_LOG_ID, DRIVE_ACTAS_ROOT

# Healthcheck
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD curl -f http://localhost:8000/api/validar || exit 1

# Arrancar via script de shell para que $PORT sea expandido correctamente
ENTRYPOINT ["/app/start.sh"]
