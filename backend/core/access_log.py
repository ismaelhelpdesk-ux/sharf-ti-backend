"""
core/access_log.py — Logs de acceso estructurados
==================================================
Registra: quién accede, cuándo, desde qué IP, qué acción,
resultado (ok/error) y duración.
Almacena en memoria (últimas 1000 entradas) + log file.
"""

import logging, time, json, asyncio
from datetime    import datetime, timezone
from typing      import Dict, List, Optional
from collections import deque
from fastapi      import Request, Response
from starlette.middleware.base import BaseHTTPMiddleware

log = logging.getLogger("sharf.access")

# ── Buffer circular de los últimos N accesos ────────────────────────────
_access_log: deque = deque(maxlen=1000)

# ════════════════════════════════════════════════════════════════════════
# MIDDLEWARE — se ejecuta en CADA request
# ════════════════════════════════════════════════════════════════════════
class AccessLogMiddleware(BaseHTTPMiddleware):
    """
    Middleware que intercepta todas las peticiones y registra:
    timestamp, IP, método, path, usuario, status, duración
    """
    SKIP_PATHS = {"/", "/docs", "/openapi.json", "/manifest.json",
                  "/static", "/favicon.ico"}

    async def dispatch(self, request: Request, call_next) -> Response:
        # Saltar rutas estáticas
        if any(request.url.path.startswith(p) for p in self.SKIP_PATHS):
            return await call_next(request)

        start = time.perf_counter()
        # IP real (considera proxies/Railway)
        ip = (request.headers.get("x-forwarded-for","").split(",")[0].strip()
              or request.headers.get("x-real-ip","")
              or (request.client.host if request.client else "unknown"))

        # Extraer usuario del JWT si viene en el header
        email = "anonymous"
        auth = request.headers.get("authorization","")
        if auth.startswith("Bearer "):
            try:
                from core.security import verificar_token
                payload = verificar_token(auth[7:])
                email = payload.get("sub", "anonymous")
            except Exception:
                pass

        response = await call_next(request)
        duration_ms = round((time.perf_counter() - start) * 1000, 1)

        entry = {
            "ts":      datetime.now(timezone.utc).isoformat(),
            "ip":      ip,
            "user":    email,
            "method":  request.method,
            "path":    request.url.path,
            "status":  response.status_code,
            "ms":      duration_ms,
            "ua":      request.headers.get("user-agent","")[:80],
        }
        _access_log.append(entry)

        # Log estructurado
        level = logging.WARNING if response.status_code >= 400 else logging.INFO
        log.log(level,
            f"{ip:15s} | {email:35s} | {request.method:6s} {request.url.path:40s} "
            f"| {response.status_code} | {duration_ms:7.1f}ms")

        return response

# ════════════════════════════════════════════════════════════════════════
# API — obtener los logs recientes
# ════════════════════════════════════════════════════════════════════════
def get_recent_logs(n: int = 100,
                    user_filter: Optional[str] = None,
                    path_filter: Optional[str] = None) -> List[Dict]:
    """Devuelve los últimos N accesos, con filtros opcionales."""
    entries = list(_access_log)[-n:]
    if user_filter:
        entries = [e for e in entries if user_filter.lower() in e["user"].lower()]
    if path_filter:
        entries = [e for e in entries if path_filter in e["path"]]
    return list(reversed(entries))   # más reciente primero

def get_stats() -> Dict:
    """Estadísticas rápidas de los últimos accesos."""
    entries = list(_access_log)
    if not entries:
        return {"total": 0}
    from collections import Counter
    users   = Counter(e["user"]   for e in entries if e["user"] != "anonymous")
    paths   = Counter(e["path"]   for e in entries)
    errors  = [e for e in entries if e["status"] >= 400]
    avg_ms  = sum(e["ms"] for e in entries[-100:]) / min(len(entries), 100)
    return {
        "total":       len(entries),
        "errors":      len(errors),
        "error_rate":  round(len(errors)/max(len(entries),1)*100, 1),
        "avg_ms":      round(avg_ms, 1),
        "top_users":   users.most_common(5),
        "top_paths":   paths.most_common(5),
        "last_access": entries[-1]["ts"] if entries else None,
    }
