"""
config.py — Configuration centralisée
Impulse AI — DocFix
"""

import os
from pathlib import Path
from typing import List


class Settings:
    # ── API ──────────────────────────────────────────────────────
    APP_NAME: str = "DocFix API — Impulse AI"
    VERSION: str  = "1.1.0"
    DEBUG: bool   = os.getenv("DEBUG", "false").lower() == "true"
    PORT: int     = int(os.getenv("PORT", "8000"))

    # ── CORS ─────────────────────────────────────────────────────
    _origins_raw: str = os.getenv(
        "ALLOWED_ORIGINS",
        "http://localhost:3000,https://localhost:3000"
    )
    ALLOWED_ORIGINS: List[str] = [o.strip() for o in _origins_raw.split(",") if o.strip()]

    # ── IA Principale : Gemini ────────────────────────────────────
    GEMINI_API_KEY: str = os.getenv("GEMINI_API_KEY", "")
    GEMINI_MODEL: str   = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")

    # ── IA Fallback (brancher ce que tu veux) ────────────────────
    # Laisser vide pour désactiver le fallback
    # Compatible : OpenAI, Mistral, Groq, Cohere, Together AI...
    # ou n'importe quelle API qui accepte le format OpenAI
    FALLBACK_AI_ENABLED: bool   = os.getenv("FALLBACK_AI_ENABLED", "false").lower() == "true"
    FALLBACK_AI_API_KEY: str    = os.getenv("FALLBACK_AI_API_KEY", "")
    FALLBACK_AI_BASE_URL: str   = os.getenv("FALLBACK_AI_BASE_URL", "")
    FALLBACK_AI_MODEL: str      = os.getenv("FALLBACK_AI_MODEL", "")
    FALLBACK_AI_NAME: str       = os.getenv("FALLBACK_AI_NAME", "Fallback AI")

    # ── Fichiers ─────────────────────────────────────────────────
    MAX_FILE_SIZE_MB: int    = int(os.getenv("MAX_FILE_SIZE_MB", "50"))
    MAX_FILE_SIZE_BYTES: int = MAX_FILE_SIZE_MB * 1024 * 1024
    JOB_EXPIRY_SECONDS: int  = int(os.getenv("JOB_EXPIRY_SECONDS", "7200"))
    TEMP_DIR: Path           = Path(os.getenv("TEMP_DIR", "/tmp/docfix"))

    # ── LibreOffice ───────────────────────────────────────────────
    LIBREOFFICE_PATH: str = os.getenv("LIBREOFFICE_PATH", "soffice")

    # ── n8n (optionnel) ───────────────────────────────────────────
    N8N_WEBHOOK_SECRET: str = os.getenv("N8N_WEBHOOK_SECRET", "")

    def validate(self):
        import logging
        log = logging.getLogger("docfix.config")
        warnings = []

        if not self.GEMINI_API_KEY:
            warnings.append("GEMINI_API_KEY manquante — analyse IA désactivée")

        if self.FALLBACK_AI_ENABLED:
            if not self.FALLBACK_AI_API_KEY:
                warnings.append("FALLBACK_AI_API_KEY manquante — fallback désactivé")
            if not self.FALLBACK_AI_BASE_URL:
                warnings.append("FALLBACK_AI_BASE_URL manquante — fallback désactivé")
            if not self.FALLBACK_AI_MODEL:
                warnings.append("FALLBACK_AI_MODEL manquant — fallback désactivé")
            else:
                log.info(f"Fallback IA activé : {self.FALLBACK_AI_NAME} ({self.FALLBACK_AI_MODEL})")

        for w in warnings:
            log.warning(f"⚠️  {w}")

        self.TEMP_DIR.mkdir(parents=True, exist_ok=True)


settings = Settings()
settings.validate()
