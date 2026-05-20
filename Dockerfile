# ── DocFix Backend — Impulse AI ──────────────────────────────────
# Image de base Python slim
FROM python:3.11-slim

# Métadonnées
LABEL org.opencontainers.image.title="DocFix API"
LABEL org.opencontainers.image.vendor="Impulse AI"
LABEL org.opencontainers.image.version="1.0.0"

# Variables d'environnement
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive \
    TEMP_DIR=/tmp/docfix \
    LIBREOFFICE_PATH=/usr/bin/soffice

# ── Dépendances système + LibreOffice ────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
    # LibreOffice headless pour la conversion PDF
    libreoffice-writer \
    libreoffice-common \
    # Polices nécessaires pour un rendu correct
    fonts-liberation \
    fonts-dejavu-core \
    fonts-noto-core \
    # Utilitaires
    curl \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# ── Dossier de travail ────────────────────────────────────────────
WORKDIR /app

# ── Dépendances Python ────────────────────────────────────────────
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

# ── Code source ──────────────────────────────────────────────────
COPY . .

# ── Dossier temporaire ────────────────────────────────────────────
RUN mkdir -p /tmp/docfix && chmod 777 /tmp/docfix

# ── Utilisateur non-root (sécurité) ─────────────────────────────
RUN useradd -m -u 1000 docfix \
    && chown -R docfix:docfix /app /tmp/docfix
USER docfix

# ── Port ─────────────────────────────────────────────────────────
EXPOSE 8000

# ── Health check ─────────────────────────────────────────────────
HEALTHCHECK --interval=30s --timeout=10s --start-period=60s --retries=3 \
    CMD curl -f http://localhost:8000/health || exit 1

# ── Démarrage ────────────────────────────────────────────────────
CMD ["uvicorn", "app.main:app", \
     "--host", "0.0.0.0", \
     "--port", "8000", \
     "--workers", "1", \
     "--timeout-keep-alive", "75"]
