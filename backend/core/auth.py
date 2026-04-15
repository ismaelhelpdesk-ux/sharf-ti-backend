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
