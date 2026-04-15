import io, base64, json, logging, asyncio
from datetime import datetime, timedelta
from typing   import Any, Dict, List, Optional
log = logging.getLogger("sharf")

def _build_google_service(creds_json: str, service: str, version: str):
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
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return build(service, version, credentials=creds, cache_discovery=False)


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
