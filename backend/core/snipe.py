"""
core/snipe.py — Cliente asíncrono para la API de Snipe-IT
Usa aiohttp para máximo rendimiento en el servidor.
"""

import aiohttp, logging, asyncio
from typing import Any, Dict, List, Optional

log = logging.getLogger("sharf.snipe")

class SnipeClient:
    def __init__(self, base_url: str, token: str):
        self.base    = base_url.rstrip("/")
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Accept":        "application/json",
            "Content-Type":  "application/json",
        }
        self._session: Optional[aiohttp.ClientSession] = None

    async def _session_get(self) -> aiohttp.ClientSession:
        if self._session is None or self._session.closed:
            timeout = aiohttp.ClientTimeout(total=15)
            self._session = aiohttp.ClientSession(
                headers=self.headers, timeout=timeout)
        return self._session

    async def _get(self, path: str, params: dict = None) -> Dict:
        s = await self._session_get()
        url = f"{self.base}/{path.lstrip('/')}"
        async with s.get(url, params=params) as r:
            r.raise_for_status()
            return await r.json()

    async def _post(self, path: str, data: dict) -> Dict:
        s = await self._session_get()
        url = f"{self.base}/{path.lstrip('/')}"
        async with s.post(url, json=data) as r:
            r.raise_for_status()
            return await r.json()

    async def _patch(self, path: str, data: dict) -> Dict:
        s = await self._session_get()
        url = f"{self.base}/{path.lstrip('/')}"
        async with s.patch(url, json=data) as r:
            r.raise_for_status()
            return await r.json()

    # ── API pública ───────────────────────────────────────────────────

    async def status(self) -> Dict:
        """Health-check: devuelve totales de hardware."""
        try:
            r = await self._get("/hardware", params={"limit": 1})
            return {"total_assets": r.get("total", 0), "ok": True}
        except Exception as e:
            log.error(f"status: {e}")
            raise

    async def get_activo(self, asset_id: int) -> Dict:
        """Obtiene detalle completo de un activo por su ID."""
        r = await self._get(f"/hardware/{asset_id}")
        return self._normalizar_activo(r)

    async def buscar_por_custom_field(self, campo: str, valor: str) -> List[Dict]:
        """
        Busca activos por campo personalizado (ej: DNI del colaborador).
        Snipe-IT permite buscar por custom fields con ?search=valor
        """
        r = await self._get("/hardware", params={
            "search": valor,
            "limit":  50,
            "status": "Deployed",   # solo activos asignados
        })
        rows = r.get("rows", [])
        # Filtrar por campo DNI en custom_fields
        resultado = []
        for row in rows:
            cfs = row.get("custom_fields", {})
            for k, cf in cfs.items():
                if campo.upper() in k.upper() and str(cf.get("value","")).strip() == str(valor).strip():
                    resultado.append(self._normalizar_activo(row))
                    break
        return resultado

    async def buscar_por_serial(self, serial: str) -> Optional[Dict]:
        """Busca un activo por número de serie."""
        r = await self._get("/hardware/byserial/" + serial)
        rows = r.get("rows", [])
        if rows:
            return self._normalizar_activo(rows[0])
        return None

    async def checkin(self, asset_id: int, datos: Dict) -> Dict:
        """
        Realiza el checkin (devolución) de un activo en Snipe-IT.
        El activo pasa a estado 'Disponible'.
        """
        payload = {
            "note":      datos.get("note", "Devolución via SHARF TI"),
            "status_id": datos.get("status_id", 4),  # 4 = Disponible
        }
        if datos.get("location_id"):
            payload["location_id"] = datos["location_id"]
        log.info(f"checkin asset_id={asset_id} payload={payload}")
        return await self._post(f"/hardware/{asset_id}/checkin", payload)

    async def get_categorias(self) -> List[Dict]:
        """Obtiene las categorías (áreas) disponibles."""
        try:
            r = await self._get("/categories", params={"limit": 100})
            return [{"id": row["id"], "nombre": row["name"]}
                    for row in r.get("rows", [])]
        except Exception as e:
            log.warning(f"get_categorias: {e}")
            return []

    async def get_locations(self) -> List[Dict]:
        """Obtiene las sedes/ubicaciones disponibles."""
        try:
            r = await self._get("/locations", params={"limit": 100})
            return [{"id": row["id"], "nombre": row["name"]}
                    for row in r.get("rows", [])]
        except Exception as e:
            log.warning(f"get_locations: {e}")
            return []

    def _normalizar_activo(self, row: Dict) -> Dict:
        """Normaliza la respuesta de Snipe-IT al formato que usa el frontend."""
        cfs = row.get("custom_fields", {})

        def cf(nombre: str) -> str:
            for k, v in cfs.items():
                if nombre.upper() in k.upper():
                    return str(v.get("value") or "")
            return ""

        return {
            "id":         row.get("id"),
            "asset_tag":  row.get("asset_tag", ""),
            "serial":     row.get("serial", ""),
            "nombre":     row.get("name", ""),
            "tipoActivo": row.get("category", {}).get("name", "Laptop") if isinstance(row.get("category"), dict) else str(row.get("category","Laptop")),
            "modelo":     row.get("model",  {}).get("name", "") if isinstance(row.get("model"),  dict) else "",
            "fabricante": row.get("manufacturer", {}).get("name","") if isinstance(row.get("manufacturer"),dict) else "",
            "estado":     row.get("status_label", {}).get("name","") if isinstance(row.get("status_label"),dict) else "",
            "ubicacion":  row.get("location", {}).get("name","") if isinstance(row.get("location"),dict) else "",
            "locationId": row.get("location", {}).get("id") if isinstance(row.get("location"),dict) else None,
            "asignado_a": row.get("assigned_to", {}).get("name","") if isinstance(row.get("assigned_to"),dict) else "",
            # Custom fields
            "dni":        cf("DNI"),
            "cargadorSerial": cf("cargador") or cf("charger"),
            "mouseDesc":  cf("mouse"),
            "mochilaDesc":cf("mochila") or cf("bag"),
            "dockingDesc":cf("docking"),
            "tecladoDesc":cf("teclado") or cf("keyboard"),
            "ceco":       cf("CECO") or cf("Centro de Costo"),
            "area":       cf("Area") or cf("Área"),
        }

    async def close(self):
        if self._session and not self._session.closed:
            await self._session.close()
