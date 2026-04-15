"""
core/security.py — JWT, autenticación y autorización
=====================================================
Implementa:
  - JWT con expiración configurable (default 1 hora)
  - Verificación de token en cada request
  - Roles: tecnico | supervisor
  - Google OAuth token validation (verificar con Google)
  - Blacklist de tokens revocados (logout)
"""

import jwt as pyjwt          # python-jose → PyJWT
import logging, time, hashlib
from datetime  import datetime, timedelta, timezone
from typing    import Optional, Dict, Set
from fastapi   import HTTPException, Security, Request
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials

from core.config import CFG

log = logging.getLogger("sharf.security")

# ── Esquema Bearer ──────────────────────────────────────────────────────
bearer = HTTPBearer(auto_error=False)

# ── Blacklist en memoria (tokens revocados por logout) ──────────────────
_revoked_tokens: Set[str] = set()

# ════════════════════════════════════════════════════════════════════════
# CREAR TOKEN JWT
# ════════════════════════════════════════════════════════════════════════
def crear_token(email: str, nombre: str, rol: str,
                expires_minutes: int = None) -> Dict:
    """
    Crea un JWT firmado con HS256.
    Devuelve: { access_token, token_type, expires_in, user }
    """
    exp_min = expires_minutes or CFG.JWT_EXPIRE_MINUTES
    now     = datetime.now(timezone.utc)
    exp     = now + timedelta(minutes=exp_min)

    payload = {
        "sub":    email,
        "nombre": nombre,
        "rol":    rol,
        "iat":    int(now.timestamp()),
        "exp":    int(exp.timestamp()),
        "jti":    hashlib.sha256(f"{email}{now.timestamp()}".encode()).hexdigest()[:16],
    }
    token = pyjwt.encode(payload, CFG.SECRET_KEY, algorithm="HS256")
    log.info(f"Token creado: {email} ({rol}) — expira en {exp_min}m")
    return {
        "access_token": token,
        "token_type":   "Bearer",
        "expires_in":   exp_min * 60,
        "expires_at":   exp.isoformat(),
        "user": {"email": email, "nombre": nombre, "rol": rol}
    }

# ════════════════════════════════════════════════════════════════════════
# VERIFICAR TOKEN JWT
# ════════════════════════════════════════════════════════════════════════
def verificar_token(token: str) -> Dict:
    """
    Valida el JWT y devuelve el payload.
    Lanza HTTPException 401 si inválido/expirado.
    """
    if not token:
        raise HTTPException(401, detail="Token requerido")

    if token in _revoked_tokens:
        raise HTTPException(401, detail="Token revocado — inicia sesión nuevamente")

    try:
        payload = pyjwt.decode(token, CFG.SECRET_KEY, algorithms=["HS256"])
        return payload
    except pyjwt.ExpiredSignatureError:
        raise HTTPException(401, detail="Token expirado — inicia sesión nuevamente")
    except pyjwt.InvalidTokenError as e:
        raise HTTPException(401, detail=f"Token inválido: {e}")

# ════════════════════════════════════════════════════════════════════════
# DEPENDENCIAS FastAPI — inyectar en endpoints
# ════════════════════════════════════════════════════════════════════════
def get_current_user(
    credentials: Optional[HTTPAuthorizationCredentials] = Security(bearer)
) -> Dict:
    """
    Dependencia: extrae y verifica el token del header Authorization.
    Uso: user = Depends(get_current_user)
    """
    if not credentials:
        raise HTTPException(401, detail="Authorization header requerido")
    return verificar_token(credentials.credentials)

def get_supervisor(
    credentials: Optional[HTTPAuthorizationCredentials] = Security(bearer)
) -> Dict:
    """
    Dependencia: verifica token Y que el usuario sea supervisor.
    Uso en endpoints de auditoría.
    """
    user = get_current_user(credentials)
    if user.get("rol") != "supervisor":
        raise HTTPException(403, detail="Acceso solo para supervisores")
    return user

def revocar_token(token: str):
    """Agrega token a la blacklist (logout)."""
    _revoked_tokens.add(token)
    log.info(f"Token revocado: {token[:20]}...")

# ════════════════════════════════════════════════════════════════════════
# GOOGLE OAUTH — verificar id_token de Google
# ════════════════════════════════════════════════════════════════════════
async def verificar_google_token(id_token: str) -> Dict:
    """
    Verifica un id_token de Google OAuth 2.0 con la API de Google.
    Devuelve el payload con email, name, picture si es válido.
    """
    import aiohttp
    try:
        async with aiohttp.ClientSession() as s:
            async with s.get(
                f"https://oauth2.googleapis.com/tokeninfo?id_token={id_token}",
                timeout=aiohttp.ClientTimeout(total=8)
            ) as r:
                if r.status != 200:
                    raise HTTPException(401, detail="Token de Google inválido")
                data = await r.json()

        email = data.get("email", "")
        if not email.endswith("@holasharf.com"):
            raise HTTPException(403, detail=f"Email {email} no pertenece a holasharf.com")

        aud = data.get("aud", "")
        if CFG.GOOGLE_CLIENT_ID and aud != CFG.GOOGLE_CLIENT_ID:
            raise HTTPException(401, detail="Client ID no coincide")

        return {
            "email":   email,
            "nombre":  data.get("name",    data.get("email","").split("@")[0]),
            "picture": data.get("picture", ""),
            "google_sub": data.get("sub", ""),
        }
    except HTTPException:
        raise
    except Exception as e:
        log.error(f"verificar_google_token: {e}")
        raise HTTPException(401, detail="No se pudo verificar el token de Google")
