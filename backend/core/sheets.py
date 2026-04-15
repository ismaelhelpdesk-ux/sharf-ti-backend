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
