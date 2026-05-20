"""
config.py — Configuration centralisée via variables d'environnement
"""

import os
from pathlib import Path
from typing import List


class Settings:
    # ── API ──────────────────────────────────────────────────────
    APP_NAME: str        = "DocFix API — Impulse AI"
    VERSION: str         = "1.0.0"
    DEBUG: bool          = os.getenv("DEBUG", "false").lower() == "true"
    PORT: int            = int(os.getenv("PORT", "8000"))

    # ── CORS ─────────────────────────────────────────────────────
    # Liste séparée par des virgules dans la variable d'env
    _origins_raw: str = os.getenv(
        "ALLOWED_ORIGINS",
        "http://localhost:3000,https://localhost:3000"
    )
    ALLOWED_ORIGINS: List[str] = [o.strip() for o in _origins_raw.split(",") if o.strip()]

    # ── Gemini ───────────────────────────────────────────────────
    GEMINI_API_KEY: str  = os.getenv("GEMINI_API_KEY", "")
    GEMINI_MODEL: str    = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")

    # ── Fichiers ─────────────────────────────────────────────────
    MAX_FILE_SIZE_MB: int    = int(os.getenv("MAX_FILE_SIZE_MB", "50"))
    MAX_FILE_SIZE_BYTES: int = MAX_FILE_SIZE_MB * 1024 * 1024
    JOB_EXPIRY_SECONDS: int  = int(os.getenv("JOB_EXPIRY_SECONDS", "7200"))  # 2h

    # Dossier temporaire (Render utilise /tmp)
    TEMP_DIR: Path = Path(os.getenv("TEMP_DIR", "/tmp/docfix"))

    # ── LibreOffice ───────────────────────────────────────────────
    # Chemin vers soffice sur Render (après apt install)
    LIBREOFFICE_PATH: str = os.getenv("LIBREOFFICE_PATH", "soffice")

    # ── n8n (optionnel) ───────────────────────────────────────────
    N8N_WEBHOOK_SECRET: str = os.getenv("N8N_WEBHOOK_SECRET", "")

    def __post_init__(self):
        self.TEMP_DIR.mkdir(parents=True, exist_ok=True)

    def validate(self):
        """Vérifie les variables critiques au démarrage."""
        errors = []
        if not self.GEMINI_API_KEY:
            errors.append("GEMINI_API_KEY manquante")
        if errors:
            import logging
            logging.getLogger("docfix").warning(
                f"Configuration incomplète : {', '.join(errors)}"
            )
        # Créer le dossier temp
        self.TEMP_DIR.mkdir(parents=True, exist_ok=True)


settings = Settings()
settings.validate()
