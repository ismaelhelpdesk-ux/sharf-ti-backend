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

    async def obtener_imagen_b64(self, nombre_archivo: str) -> str:
        """Obtiene imagen de Drive como data URL base64."""
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(None, self._img_b64_sync, nombre_archivo)

    def _img_b64_sync(self, nombre: str) -> str:
        svc = self._svc_get()
        if not svc: return ""
        from core.config import CFG
        folder = CFG.DRIVE_FIRMAS_ID or self.root
        q = f"name=\'{nombre}\' and \'{folder}\' in parents and trashed=false"
        r = svc.files().list(q=q, fields="files(id,name,mimeType)").execute()
        files = r.get("files", [])
        if not files:
            q2 = f"name=\'{nombre}\' and trashed=false"
            r2 = svc.files().list(q=q2, fields="files(id,name,mimeType)").execute()
            files = r2.get("files", [])
        if not files: return ""
        fid  = files[0]["id"]
        mime = files[0].get("mimeType", "image/png")
        data = svc.files().get_media(fileId=fid).execute()
        b64  = base64.b64encode(data).decode()
        return f"data:{mime};base64,{b64}"

    async def buscar_acta_asignacion(self, nombre: str, dni: str, serial: str) -> dict:
        """Busca acta de asignación original en Drive."""
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(None, self._buscar_acta_sync, nombre, dni, serial)

    def _buscar_acta_sync(self, nombre: str, dni: str, serial: str) -> dict:
        svc = self._svc_get()
        if not svc: return {"error": "Drive sin credenciales"}
        terminos = [t for t in [serial, nombre.split(",")[0].strip() if nombre else ""] if t]
        archivos = []
        for term in terminos[:2]:
            q = f"name contains \'{term}\' and \'{self.root}\' in parents and mimeType=\'application/pdf\' and trashed=false"
            try:
                r = svc.files().list(q=q, fields="files(id,name,webViewLink,createdTime)",
                                     orderBy="createdTime desc", pageSize=5).execute()
                for f in r.get("files", []):
                    if f not in archivos: archivos.append(f)
            except Exception as e:
                log.warning(f"buscar_acta({term}): {e}")
        if not archivos:
            return {"archivos":[], "carpetaUrl":"",
                    "error":f"No se encontraron actas para {nombre or serial}"}
        return {"archivos":[{"nombre":f["name"],"url":f["webViewLink"],
                             "fecha":f.get("createdTime","")} for f in archivos[:5]],
                "carpetaUrl":f"https://drive.google.com/drive/folders/{self.root}",
                "carpetaNombre":nombre}

