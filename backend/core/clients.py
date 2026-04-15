"""
core/sheets.py — Google Sheets (historial de devoluciones)
core/drive.py  — Google Drive (subir PDFs)
core/mail.py   — Gmail (enviar correos con adjunto)
core/qr.py     — Decodificar QR con pyzbar
core/auth.py   — Autenticación de técnicos
"""

# ══════════════════════════════════════════════════════════════════════════
# SHEETS
# ══════════════════════════════════════════════════════════════════════════
import io, base64, json, logging, asyncio
from datetime import datetime, timedelta
from typing   import Any, Dict, List, Optional
from functools import lru_cache

log = logging.getLogger("sharf")

def _build_google_service(creds_json: str, service: str, version: str):
    """Construye un cliente de Google API a partir del JSON de Service Account."""
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/gmail.send",
    ]
    if not creds_json:
        raise ValueError("GOOGLE_CREDS_JSON no configurado")
    info = json.loads(creds_json) if creds_json.strip().startswith("{") else json.load(open(creds_json))
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return build(service, version, credentials=creds, cache_discovery=False)


class SheetsClient:
    def __init__(self, sheet_id: str, creds_json: str):
        self.sheet_id  = sheet_id
        self.creds_json = creds_json
        self._svc      = None

    def _svc_get(self):
        if self._svc is None and self.creds_json:
            self._svc = _build_google_service(self.creds_json, "sheets", "v4")
        return self._svc

    async def contar_filas(self) -> int:
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(None, self._contar_filas_sync)

    def _contar_filas_sync(self) -> int:
        svc = self._svc_get()
        if not svc:
            return -1
        result = svc.spreadsheets().values().get(
            spreadsheetId=self.sheet_id, range="A:A").execute()
        vals = result.get("values", [])
        return max(0, len(vals) - 1)

    async def registrar_devolucion(self, p, dev_id: str,
                                   drive_url: str, snipe_ok: bool, pdf_ok: bool):
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(None, self._registrar_sync,
                                   p, dev_id, drive_url, snipe_ok, pdf_ok)

    def _registrar_sync(self, p, dev_id, drive_url, snipe_ok, pdf_ok):
        svc = self._svc_get()
        if not svc:
            raise RuntimeError("Sheets no disponible (sin credenciales)")
        emp = p.empleado or {}
        tec = p.tecnico  or {}
        ts  = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        row = [
            dev_id, ts, p.tipoDev, p.modoIngreso,
            emp.get("nombre",""), emp.get("dni",""), emp.get("empresa",""),
            emp.get("area",""), emp.get("cargo",""), emp.get("sede",""),
            p.tipoActivo, p.serial,
            str(p.snipeId or ""),
            ",".join(k for k,v in (p.accesoriosDevueltos or {}).items() if v),
            ",".join(k for k,v in (p.accesoriosCotizar  or {}).items() if v),
            sum((p.accesoriosCosto or {}).values()),
            "BUENO" if p.equipoBueno else "OBSERVACIONES",
            p.observacionDesc or "",
            tec.get("nombre",""), tec.get("email",""), tec.get("sede",""),
            "✅" if snipe_ok else "❌",
            "✅" if pdf_ok   else "❌",
            drive_url or "",
        ]
        svc.spreadsheets().values().append(
            spreadsheetId=self.sheet_id,
            range="Devoluciones_TI!A:Z",
            valueInputOption="USER_ENTERED",
            body={"values": [row]}
        ).execute()
        log.info(f"Sheets: fila registrada {dev_id}")

    async def get_auditoria(self, meses: int) -> Dict:
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(None, self._auditoria_sync, meses)

    def _auditoria_sync(self, meses: int) -> Dict:
        svc = self._svc_get()
        if not svc:
            return {"vacio": True, "error": "Sin credenciales"}
        result = svc.spreadsheets().values().get(
            spreadsheetId=self.sheet_id, range="Devoluciones_TI!A:Z").execute()
        rows = result.get("values", [])
        if len(rows) < 2:
            return {"vacio": True, "filas": 0}

        hdrs = rows[0]
        data = rows[1:]
        limite = datetime.now() - timedelta(days=meses*30)

        # Índices
        def col(nombre):
            for i,h in enumerate(hdrs):
                if nombre.lower() in h.lower(): return i
            return -1
        iDate = col("Fecha"); iTipo = col("Tipo Dev"); iTec = col("Técnico")
        iArea = col("Área"); iSnipe = col("Snipe"); iEmail = col("Email")

        filas_periodo = []
        for r in data:
            try:
                ts = datetime.strptime(r[iDate] if iDate>=0 else "", "%d/%m/%Y %H:%M:%S")
                if ts >= limite: filas_periodo.append(r)
            except: pass

        # KPIs
        total   = len(filas_periodo)
        por_tec = {}; por_tipo = {}; por_mes = {}
        err_snipe = 0; err_email = 0

        for r in filas_periodo:
            tec   = r[iTec]   if iTec>=0   and len(r)>iTec   else "?"
            tipo  = r[iTipo]  if iTipo>=0  and len(r)>iTipo  else "?"
            mes   = r[iDate][:7] if iDate>=0 and len(r)>iDate else "?"
            snipe = r[iSnipe] if iSnipe>=0 and len(r)>iSnipe else ""
            email = r[iEmail] if iEmail>=0 and len(r)>iEmail else ""
            por_tec[tec]  = por_tec.get(tec, 0) + 1
            por_tipo[tipo] = por_tipo.get(tipo, 0) + 1
            por_mes[mes]   = por_mes.get(mes, 0) + 1
            if "❌" in snipe: err_snipe += 1
            if "❌" in email: err_email += 1

        return {
            "vacio": total == 0, "filas": total,
            "periodo": f"Últimos {meses} meses",
            "resumen": {"total": total, "errSnipe": err_snipe, "errEmail": err_email,
                        "pctObs": 0, "conObs": 0, "totalNoDev": 0, "totalCotiz": 0},
            "porTecnico": [{"label":k,"val":v} for k,v in sorted(por_tec.items(), key=lambda x:-x[1])],
            "porTipo":    [{"label":k,"val":v} for k,v in sorted(por_tipo.items(), key=lambda x:-x[1])],
            "porMes":     [{"mes":k,"total":v} for k,v in sorted(por_mes.items())],
        }


# ══════════════════════════════════════════════════════════════════════════
# DRIVE
# ══════════════════════════════════════════════════════════════════════════
from googleapiclient.http import MediaIoBaseUpload

class DriveClient:
    def __init__(self, root_folder_id: str, creds_json: str):
        self.root       = root_folder_id
        self.creds_json = creds_json
        self._svc       = None

    def _svc_get(self):
        if self._svc is None and self.creds_json:
            self._svc = _build_google_service(self.creds_json, "drive", "v3")
        return self._svc

    async def verificar_carpeta(self) -> str:
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(None, self._verificar_sync)

    def _verificar_sync(self) -> str:
        svc = self._svc_get()
        if not svc: raise RuntimeError("Drive sin credenciales")
        r = svc.files().get(fileId=self.root, fields="name").execute()
        return r.get("name","OK")

    async def subir_pdf(self, pdf_bytes: Optional[bytes], nombre: str,
                        empleado: str = "", tecnico: str = "") -> str:
        if not pdf_bytes:
            return ""
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(None, self._subir_sync,
                                          pdf_bytes, nombre, empleado, tecnico)

    def _subir_sync(self, pdf_bytes, nombre, empleado, tecnico) -> str:
        svc = self._svc_get()
        if not svc: raise RuntimeError("Drive sin credenciales")

        # Crear/encontrar subcarpeta del técnico
        folder_id = self._get_or_create_folder(svc, tecnico or "Sin técnico", self.root)

        meta = {"name": nombre, "parents": [folder_id],
                "mimeType": "application/pdf"}
        media = MediaIoBaseUpload(io.BytesIO(pdf_bytes),
                                  mimetype="application/pdf", resumable=False)
        f = svc.files().create(body=meta, media_body=media,
                               fields="id,webViewLink").execute()
        # Hacer público el enlace
        svc.permissions().create(
            fileId=f["id"],
            body={"type":"anyone","role":"reader"}
        ).execute()
        url = f.get("webViewLink", f"https://drive.google.com/file/d/{f['id']}/view")
        log.info(f"Drive: {nombre} → {url}")
        return url

    def _get_or_create_folder(self, svc, nombre: str, parent_id: str) -> str:
        q = (f"name='{nombre}' and '{parent_id}' in parents "
             f"and mimeType='application/vnd.google-apps.folder' and trashed=false")
        r = svc.files().list(q=q, fields="files(id)").execute()
        files = r.get("files", [])
        if files:
            return files[0]["id"]
        meta = {"name": nombre, "parents": [parent_id],
                "mimeType": "application/vnd.google-apps.folder"}
        f = svc.files().create(body=meta, fields="id").execute()
        return f["id"]


# ══════════════════════════════════════════════════════════════════════════
# MAIL
# ══════════════════════════════════════════════════════════════════════════
import email as email_lib
from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText
from email.mime.base      import MIMEBase
from email               import encoders

class MailClient:
    def __init__(self, usuario: str, creds_json: str):
        self.usuario    = usuario
        self.creds_json = creds_json
        self._svc       = None

    def _svc_get(self):
        if self._svc is None and self.creds_json:
            self._svc = _build_google_service(self.creds_json, "gmail", "v1")
        return self._svc

    async def verificar(self):
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(None, self._verificar_sync)

    def _verificar_sync(self):
        svc = self._svc_get()
        if not svc: raise RuntimeError("Gmail sin credenciales")
        svc.users().getProfile(userId="me").execute()

    async def enviar_devolucion(self, p, dev_id: str,
                                pdf_bytes: Optional[bytes], drive_url: str):
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(None, self._enviar_sync,
                                   p, dev_id, pdf_bytes, drive_url)

    def _enviar_sync(self, p, dev_id, pdf_bytes, drive_url):
        svc = self._svc_get()
        if not svc: raise RuntimeError("Gmail sin credenciales")

        emp = p.empleado or {}
        tec = p.tecnico  or {}
        destinatario = p.emailPara or emp.get("emailPers","") or emp.get("emailCorp","")
        cc  = p.emailCC  or f"helpdesk@holasharf.com,anais.chero@holasharf.com"
        bcc = p.emailBCC or ""

        asunto = (f"Acta de Devolución de Equipos — {emp.get('nombre','Colaborador')} "
                  f"[{dev_id}]")
        cuerpo = f"""
<html><body style="font-family:Arial,sans-serif;color:#1a1a1a">
<div style="max-width:600px;margin:0 auto">
  <div style="background:#9B1035;padding:20px;text-align:center">
    <h1 style="color:#fff;margin:0;font-style:italic">sharf</h1>
    <p style="color:#FFC1C2;margin:4px 0 0;font-size:12px">Mesa de Servicio TI</p>
  </div>
  <div style="padding:24px;background:#fdf5f6;border:1px solid #f0d6d8">
    <h2 style="color:#9B1035;margin-top:0">Devolución de Equipos registrada</h2>
    <p>Estimado/a <b>{emp.get('nombre','colaborador')}</b>,</p>
    <p>Se ha registrado la devolución de sus equipos tecnológicos.</p>
    <table style="width:100%;border-collapse:collapse;margin:16px 0">
      <tr style="background:#FF6568;color:#fff">
        <th style="padding:8px;text-align:left">Campo</th>
        <th style="padding:8px;text-align:left">Detalle</th>
      </tr>
      <tr style="background:#fff"><td style="padding:8px;border:1px solid #f0d6d8"><b>ID Devolución</b></td><td style="padding:8px;border:1px solid #f0d6d8">{dev_id}</td></tr>
      <tr style="background:#fdf5f6"><td style="padding:8px;border:1px solid #f0d6d8"><b>Tipo</b></td><td style="padding:8px;border:1px solid #f0d6d8">{p.tipoDev}</td></tr>
      <tr style="background:#fff"><td style="padding:8px;border:1px solid #f0d6d8"><b>Activo</b></td><td style="padding:8px;border:1px solid #f0d6d8">{p.tipoActivo} — S/N: {p.serial}</td></tr>
      <tr style="background:#fdf5f6"><td style="padding:8px;border:1px solid #f0d6d8"><b>Técnico</b></td><td style="padding:8px;border:1px solid #f0d6d8">{tec.get('nombre','')}</td></tr>
      <tr style="background:#fff"><td style="padding:8px;border:1px solid #f0d6d8"><b>Fecha</b></td><td style="padding:8px;border:1px solid #f0d6d8">{datetime.now().strftime('%d/%m/%Y %H:%M')}</td></tr>
    </table>
    {'<p><a href="' + drive_url + '" style="background:#9B1035;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;font-weight:bold">📄 Ver Acta en Google Drive</a></p>' if drive_url else ''}
    <p style="font-size:12px;color:#b08090;margin-top:16px">
      El acta firmada se adjunta a este correo y también está disponible en Google Drive.
    </p>
  </div>
  <div style="background:#f0d6d8;padding:12px;text-align:center;font-size:11px;color:#5c3040">
    SHARF Mesa de Servicio TI · helpdesk@holasharf.com
  </div>
</div>
</body></html>"""

        msg = MIMEMultipart("mixed")
        msg["From"]    = self.usuario
        msg["To"]      = destinatario
        if cc:  msg["Cc"]  = cc
        if bcc: msg["Bcc"] = bcc
        msg["Subject"] = asunto
        msg.attach(MIMEText(cuerpo, "html"))

        if pdf_bytes:
            part = MIMEBase("application","octet-stream")
            part.set_payload(pdf_bytes)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f'attachment; filename="Acta_{dev_id}.pdf"')
            msg.attach(part)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        svc.users().messages().send(userId="me", body={"raw": raw}).execute()
        log.info(f"Mail: enviado a {destinatario}")


# ══════════════════════════════════════════════════════════════════════════
# QR DECODER — pyzbar (más potente que jsQR)
# ══════════════════════════════════════════════════════════════════════════
def decodificar_qr(imagen_bytes: bytes) -> Optional[Dict]:
    """
    Decodifica un QR desde bytes de imagen usando pyzbar.
    Parsea el formato SHARF: SERIAL=... - DNI=... - NOMBRE=... - CECO=...
    """
    try:
        from PIL    import Image as PILImage
        from pyzbar import pyzbar
        img  = PILImage.open(io.BytesIO(imagen_bytes)).convert("RGB")
        # Intentar múltiples escalas para mejorar la detección
        for scale in [1.0, 1.5, 2.0]:
            if scale != 1.0:
                w, h = img.size
                img_s = img.resize((int(w*scale), int(h*scale)))
            else:
                img_s = img
            qrs = pyzbar.decode(img_s)
            if qrs:
                texto = qrs[0].data.decode("utf-8", errors="ignore").strip()
                log.info(f"QR decodificado: {texto[:80]}")
                return _parsear_qr(texto)
        return None
    except ImportError:
        log.warning("pyzbar no disponible — usando fallback")
        return None
    except Exception as e:
        log.error(f"decodificar_qr: {e}")
        return None

def _parsear_qr(texto: str) -> Dict:
    """
    Parsea el formato QR de SHARF:
    SERIAL=SN8X4K2LM - DNI=45123678 - NOMBRE=QUISPE LIMA Maria - CECO=PE001 - CARGADOR=CHG-001
    """
    resultado = {"raw": texto}
    partes = [p.strip() for p in texto.replace(" - ","\n").replace("-\n","\n").split("\n")]
    for parte in partes:
        if "=" in parte:
            k, _, v = parte.partition("=")
            k = k.strip().upper()
            v = v.strip()
            if   k == "SERIAL":   resultado["serial"]   = v
            elif k == "DNI":      resultado["dni"]      = v
            elif k == "NOMBRE":   resultado["nombre"]   = v
            elif k == "CECO":     resultado["ceco"]     = v
            elif k == "CARGADOR": resultado["cargador"] = v
            elif k == "AREA":     resultado["area"]     = v
    return resultado


# ══════════════════════════════════════════════════════════════════════════
# AUTH — lista de técnicos autorizados
# ══════════════════════════════════════════════════════════════════════════
USUARIOS_AUTORIZADOS = {
    "ismael.helpdesk@holasharf.com":  {"nombre":"Ismael Gomez Sime",    "sede":"Callao","firmaKey":"firma_ismael",   "rol":"supervisor"},
    "anais.chero@holasharf.com":      {"nombre":"Anais Chero Benites",  "sede":"Callao","firmaKey":"firma_anais",    "rol":"supervisor"},
    "eddie.fernandez@holasharf.com":  {"nombre":"Eddie Fernandez",      "sede":"Callao","firmaKey":"firma_eddie",    "rol":"supervisor"},
    "jesus.helpdesk@holasharf.com":   {"nombre":"Jesus Helpdesk",       "sede":"Callao","firmaKey":"firma_jesus",    "rol":"tecnico"},
    "michael.helpdesk@holasharf.com": {"nombre":"Michael Helpdesk",     "sede":"Callao","firmaKey":"firma_michael",  "rol":"tecnico"},
    "gabriel.helpdesk@holasharf.com": {"nombre":"Gabriel García",       "sede":"Callao","firmaKey":"firma_gabriel",  "rol":"tecnico"},
    "rocio.helpdesk@holasharf.com":   {"nombre":"Rocio Helpdesk",       "sede":"Callao","firmaKey":"firma_rocio",    "rol":"tecnico"},
    "perla.helpdesk@holasharf.com":   {"nombre":"Perla Moreno",         "sede":"Paita", "firmaKey":"firma_perla",    "rol":"tecnico"},
    "alfredo.helpdesk@holasharf.com": {"nombre":"Alfredo Helpdesk",     "sede":"Callao","firmaKey":"firma_alfredo",  "rol":"tecnico"},
    "misael.helpdesk@holasharf.com":  {"nombre":"Misael (Victor G.)",   "sede":"Callao","firmaKey":"firma_misael",   "rol":"tecnico"},
}

def autenticar_tecnico(email: str) -> Optional[Dict]:
    """Retorna datos del técnico si está autorizado, None si no."""
    tec = USUARIOS_AUTORIZADOS.get(email.strip().lower())
    if tec:
        return {"email": email, **tec}
    return None
