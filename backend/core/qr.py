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
