"""
DocFix — Backend FastAPI
Développé par Impulse AI
Déployé sur Render
"""

import os
import uuid
import asyncio
import logging
from contextlib import asynccontextmanager
from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from typing import Optional

from .jobs import job_store
from .processor import process_document
from .config import settings

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
)
logger = logging.getLogger("docfix")


# ── Nettoyage périodique des jobs expirés ────────────────────────────────────
@asynccontextmanager
async def lifespan(app: FastAPI):
    # Démarrage
    logger.info("DocFix API démarrée — Impulse AI")
    cleanup_task = asyncio.create_task(periodic_cleanup())
    yield
    # Arrêt
    cleanup_task.cancel()
    logger.info("DocFix API arrêtée")


async def periodic_cleanup():
    """Supprime les jobs et fichiers temporaires toutes les 30 min."""
    while True:
        await asyncio.sleep(1800)
        removed = job_store.cleanup_expired()
        if removed:
            logger.info(f"Nettoyage : {removed} jobs expirés supprimés")


# ── Application ───────────────────────────────────────────────────────────────
app = FastAPI(
    title="DocFix API — Impulse AI",
    description="API de correction automatique de documents Word",
    version="1.0.0",
    lifespan=lifespan,
)

# ── CORS ──────────────────────────────────────────────────────────────────────
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


# ── Schémas Pydantic ──────────────────────────────────────────────────────────
class StepSchema(BaseModel):
    id: str
    label: str
    status: str  # pending | running | done | error


class JobStatusSchema(BaseModel):
    jobId: str
    status: str           # queued | processing | done | error
    progress: int         # 0-100
    currentStep: str
    steps: list[StepSchema]
    error: Optional[str] = None
    result: Optional[dict] = None


# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/", tags=["Health"])
async def root():
    return {
        "service": "DocFix API",
        "vendor": "Impulse AI",
        "status": "online",
        "version": "1.0.0",
    }


@app.get("/health", tags=["Health"])
async def health():
    return {"status": "ok"}


@app.post("/api/v1/process", tags=["Processing"])
async def process(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
):
    """
    Reçoit un fichier .docx, crée un job et lance le traitement en arrière-plan.
    Retourne immédiatement un job_id pour le polling.
    """
    # Validation
    if not file.filename:
        raise HTTPException(400, "Nom de fichier manquant")

    fname = file.filename.lower()
    if not (fname.endswith(".docx") or fname.endswith(".doc")):
        raise HTTPException(400, "Seuls les fichiers .docx et .doc sont acceptés")

    content = await file.read()
    if len(content) > settings.MAX_FILE_SIZE_BYTES:
        raise HTTPException(413, f"Fichier trop volumineux (max {settings.MAX_FILE_SIZE_MB} Mo)")
    if len(content) < 100:
        raise HTTPException(400, "Fichier vide ou corrompu")

    # Créer le job
    job_id = str(uuid.uuid4())
    job_store.create(job_id, file.filename)

    # Sauvegarder le fichier temporairement
    input_path = settings.TEMP_DIR / f"{job_id}_input.docx"
    input_path.write_bytes(content)

    # Lancer le traitement en arrière-plan
    background_tasks.add_task(process_document, job_id, input_path)

    logger.info(f"Job créé : {job_id} — fichier : {file.filename} ({len(content)//1024} Ko)")

    return JSONResponse(
        status_code=202,
        content={"job_id": job_id, "status": "queued"},
    )


@app.post("/api/v1/webhook/process", tags=["Webhook n8n"])
async def webhook_process(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
):
    """
    Endpoint webhook compatible n8n.
    Même logique que /api/v1/process — point d'entrée alternatif.
    """
    return await process(background_tasks, file)


@app.get("/api/v1/status/{job_id}", response_model=JobStatusSchema, tags=["Processing"])
async def get_status(job_id: str):
    """Retourne l'état courant d'un job (pour polling côté frontend)."""
    job = job_store.get(job_id)
    if not job:
        raise HTTPException(404, f"Job {job_id} introuvable ou expiré")
    return job.to_schema()


@app.get("/api/v1/download/{job_id}/{format}", tags=["Download"])
async def download(job_id: str, format: str):
    """
    Télécharge le fichier résultat (docx ou pdf).
    Le fichier est supprimé du serveur 2h après sa création.
    """
    if format not in ("docx", "pdf"):
        raise HTTPException(400, "Format invalide. Utilisez 'docx' ou 'pdf'")

    job = job_store.get(job_id)
    if not job:
        raise HTTPException(404, "Job introuvable ou expiré")
    if job.status != "done":
        raise HTTPException(409, f"Job en cours ou en erreur (status: {job.status})")

    output_path = settings.TEMP_DIR / f"{job_id}_output.{format}"
    if not output_path.exists():
        raise HTTPException(404, f"Fichier {format.upper()} introuvable sur le serveur")

    media_types = {
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pdf":  "application/pdf",
    }
    original_name = job.filename.rsplit(".", 1)[0]
    download_name = f"{original_name}_corrige.{format}"

    return FileResponse(
        path=str(output_path),
        media_type=media_types[format],
        filename=download_name,
        headers={"Content-Disposition": f'attachment; filename="{download_name}"'},
    )


# ── Gestionnaire d'erreurs global ─────────────────────────────────────────────
@app.exception_handler(Exception)
async def global_exception_handler(request, exc):
    logger.error(f"Erreur non gérée : {exc}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"detail": "Erreur interne du serveur", "vendor": "Impulse AI DocFix"},
    )
