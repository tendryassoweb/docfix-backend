"""
config.py v2.0 — avec N8N_GEMINI_WEBHOOK_URL
Impulse AI — DocFix
"""

import os
from pathlib import Path
from typing import List


class Settings:
    APP_NAME: str = "DocFix API — Impulse AI"
    VERSION: str  = "2.0.0"
    DEBUG: bool   = os.getenv("DEBUG", "false").lower() == "true"
    PORT: int     = int(os.getenv("PORT", "8000"))

    # ── CORS ─────────────────────────────────────────────────────
    _origins_raw: str = os.getenv(
        "ALLOWED_ORIGINS",
        "http://localhost:3000,https://localhost:3000"
    )
    ALLOWED_ORIGINS: List[str] = [
        o.strip() for o in _origins_raw.split(",") if o.strip()
    ]

    # ── IA via n8n webhook ────────────────────────────────────────
    # URL du webhook n8n qui appelle Gemini
    # Ex: https://ton-n8n.onrender.com/webhook/gemini-analyze
    N8N_GEMINI_WEBHOOK_URL: str = os.getenv("N8N_GEMINI_WEBHOOK_URL", "")

    # Secret optionnel pour sécuriser le webhook
    N8N_WEBHOOK_SECRET: str = os.getenv("N8N_WEBHOOK_SECRET", "")

    # ── Fichiers ─────────────────────────────────────────────────
    MAX_FILE_SIZE_MB: int    = int(os.getenv("MAX_FILE_SIZE_MB", "50"))
    MAX_FILE_SIZE_BYTES: int = MAX_FILE_SIZE_MB * 1024 * 1024
    JOB_EXPIRY_SECONDS: int  = int(os.getenv("JOB_EXPIRY_SECONDS", "7200"))
    TEMP_DIR: Path           = Path(os.getenv("TEMP_DIR", "/tmp/docfix"))

    # ── LibreOffice ───────────────────────────────────────────────
    LIBREOFFICE_PATH: str = os.getenv("LIBREOFFICE_PATH", "soffice")

    # ── Fallback (désactivé — géré par n8n maintenant) ───────────
    FALLBACK_AI_ENABLED: bool = False

    def validate(self):
        import logging
        log = logging.getLogger("docfix.config")
        self.TEMP_DIR.mkdir(parents=True, exist_ok=True)

        if self.N8N_GEMINI_WEBHOOK_URL:
            log.info(f"✅ n8n webhook configuré : {self.N8N_GEMINI_WEBHOOK_URL}")
        else:
            log.warning("⚠️  N8N_GEMINI_WEBHOOK_URL non configurée — heuristique locale uniquement")


settings = Settings()
settings.validate()
