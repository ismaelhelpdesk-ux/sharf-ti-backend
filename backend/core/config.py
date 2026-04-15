"""
core/config.py — Configuración centralizada del servidor SHARF
Todas las credenciales se leen de variables de entorno.
"""
import os
from dataclasses import dataclass, field
from typing import List

@dataclass
class Config:
    # ── Snipe-IT ───────────────────────────────────────────────────────
    SNIPE_BASE:  str = os.getenv("SNIPE_BASE",  "https://scharff.snipe-it.io/api/v1")
    SNIPE_TOKEN: str = os.getenv("SNIPE_TOKEN", "")

    # ── Google ─────────────────────────────────────────────────────────
    GOOGLE_CREDS_JSON: str = os.getenv("GOOGLE_CREDS_JSON", "")
    SHEET_LOG_ID:      str = os.getenv("SHEET_LOG_ID",      "14LNid3_E7deg9rHD65TA9P_h8Lmr2s0HvWoejHlYQow")
    DRIVE_ACTAS_ROOT:  str = os.getenv("DRIVE_ACTAS_ROOT",  "0APknu_tBOg5SUk9PVA")
    DRIVE_FIRMAS_ID:   str = os.getenv("DRIVE_FIRMAS_ID",   "1EXUIcq56A23yMmvjco56gYwKr5mFRP0v")
    GMAIL_USER:        str = os.getenv("GMAIL_USER",        "helpdesk@holasharf.com")

    # ── Gmail — tres opciones ──────────────────────────────────────────
    GMAIL_APP_PASSWORD: str = os.getenv("GMAIL_APP_PASSWORD", "")
    GAS_WEBHOOK_URL:    str = os.getenv("GAS_WEBHOOK_URL",    "")
    GAS_WEBHOOK_TOKEN:  str = os.getenv("GAS_WEBHOOK_TOKEN",  "sharf-secret")

    # ── JWT ────────────────────────────────────────────────────────────
    SECRET_KEY:          str = os.getenv("SECRET_KEY", "sharf-dev-secret-change-in-prod-32chars")
    JWT_EXPIRE_MINUTES:  int = int(os.getenv("JWT_EXPIRE_MINUTES",  "60"))
    JWT_REFRESH_MINUTES: int = int(os.getenv("JWT_REFRESH_MINUTES", "1440"))

    # ── Google OAuth ───────────────────────────────────────────────────
    GOOGLE_CLIENT_ID: str = os.getenv("GOOGLE_CLIENT_ID", "")

    # ── Rate limiting ──────────────────────────────────────────────────
    RATE_LIMIT_PER_MINUTE: int = int(os.getenv("RATE_LIMIT_PER_MINUTE", "60"))
    RATE_LIMIT_AUTH:       int = int(os.getenv("RATE_LIMIT_AUTH",       "10"))

    # ── Seguridad ──────────────────────────────────────────────────────
    ADMIN_TOKEN: str = os.getenv("ADMIN_TOKEN", "")

    # ── CORS ───────────────────────────────────────────────────────────
    ALLOWED_ORIGINS: List[str] = field(default_factory=lambda: [
        o for o in [os.getenv("FRONTEND_URL", ""), "http://localhost:3000",
                    "http://localhost:8000"] if o
    ])

    # ── App ────────────────────────────────────────────────────────────
    ENV:       str = os.getenv("ENV", "development")
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "INFO")
    PORT:      int = int(os.getenv("PORT", "8000"))
    TZ:        str = os.getenv("TZ", "America/Lima")

    @property
    def is_production(self) -> bool:
        return self.ENV == "production"

    @property
    def mail_mode(self) -> str:
        if self.GMAIL_APP_PASSWORD: return "smtp"
        if self.GAS_WEBHOOK_URL:    return "hybrid"
        if self.GOOGLE_CREDS_JSON:  return "service_account"
        return "none"

    def validate(self):
        missing = []
        if not self.SNIPE_TOKEN:       missing.append("SNIPE_TOKEN")
        if not self.GOOGLE_CREDS_JSON: missing.append("GOOGLE_CREDS_JSON")
        if missing:
            import warnings
            warnings.warn(f"Variables faltantes: {', '.join(missing)}", RuntimeWarning, stacklevel=2)

CFG = Config()
CFG.validate()
