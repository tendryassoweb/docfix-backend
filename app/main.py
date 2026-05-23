"""
main.py v1.2
Impulse AI — DocFix

Nouveautés :
- Paramètres start_page et include_toc dans /api/v1/process
"""

import os, uuid, asyncio, logging
from contextlib import asynccontextmanager
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from typing import Optional

from .jobs import job_store
from .processor import process_document
from .ai_service import check_gemini_status, check_fallback_status
from .config import settings

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
)
logger = logging.getLogger("docfix")


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info(f"DocFix API v{settings.VERSION} — Impulse AI")
    task = asyncio.create_task(_periodic_cleanup())
    yield
    task.cancel()


async def _periodic_cleanup():
    while True:
        await asyncio.sleep(1800)
        removed = job_store.cleanup_expired()
        if removed:
            logger.info(f"Nettoyage: {removed} jobs expires")


app = FastAPI(
    title="DocFix API — Impulse AI",
    description="Correction automatique de documents Word",
    version=settings.VERSION,
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


class StepSchema(BaseModel):
    id: str; label: str; status: str

class JobStatusSchema(BaseModel):
    jobId: str; status: str; progress: int; currentStep: str
    steps: list[StepSchema]; error: Optional[str] = None
    result: Optional[dict] = None


@app.get("/", tags=["Health"])
async def root():
    return {
        "service": "DocFix API",
        "vendor":  "Impulse AI",
        "version": settings.VERSION,
        "status":  "online",
    }

@app.get("/health", tags=["Health"])
async def health():
    return {"status": "ok"}

@app.get("/api/v1/ai/status", tags=["AI"])
async def ai_status():
    gemini   = await check_gemini_status()
    fallback = await check_fallback_status()
    active = "heuristic"
    if gemini["status"] == "ok":
        active = "gemini"
    elif fallback["status"] == "ok":
        active = "fallback"
    return {
        "active_provider":  active,
        "gemini":           gemini,
        "fallback":         fallback,
        "fallback_enabled": settings.FALLBACK_AI_ENABLED,
    }


@app.post("/api/v1/process", tags=["Processing"])
async def process(
    background_tasks: BackgroundTasks,
    file:         UploadFile = File(...),
    start_page:   int        = Form(default=1,    ge=1, le=9999),
    include_toc:  bool       = Form(default=True),
):
    """
    Upload + lancer le traitement.

    Paramètres de formulaire :
    - file        : fichier .docx ou .doc
    - start_page  : numéro de la première page (défaut: 1)
    - include_toc : inclure la table des matières (défaut: true)
    """
    if not file.filename:
        raise HTTPException(400, "Nom de fichier manquant")

    fname = file.filename.lower()
    if not (fname.endswith(".docx") or fname.endswith(".doc")):
        raise HTTPException(400, "Seuls les fichiers .docx et .doc sont acceptes")

    content = await file.read()
    if len(content) > settings.MAX_FILE_SIZE_BYTES:
        raise HTTPException(413, f"Fichier trop volumineux (max {settings.MAX_FILE_SIZE_MB} Mo)")
    if len(content) < 100:
        raise HTTPException(400, "Fichier vide ou corrompu")

    job_id     = str(uuid.uuid4())
    input_path = settings.TEMP_DIR / f"{job_id}_input.docx"
    input_path.write_bytes(content)
    job_store.create(job_id, file.filename)

    background_tasks.add_task(
        process_document,
        job_id,
        input_path,
        start_page,
        include_toc,
    )

    logger.info(
        f"Job {job_id} — {file.filename} "
        f"({len(content)//1024} Ko) "
        f"start_page={start_page} toc={include_toc}"
    )

    return JSONResponse(
        status_code=202,
        content={
            "job_id":      job_id,
            "status":      "queued",
            "start_page":  start_page,
            "include_toc": include_toc,
        },
    )


@app.post("/api/v1/webhook/process", tags=["Webhook"])
async def webhook_process(
    background_tasks: BackgroundTasks,
    file:        UploadFile = File(...),
    start_page:  int        = Form(default=1,   ge=1, le=9999),
    include_toc: bool       = Form(default=True),
):
    return await process(background_tasks, file, start_page, include_toc)


@app.get("/api/v1/status/{job_id}", response_model=JobStatusSchema, tags=["Processing"])
async def get_status(job_id: str):
    job = job_store.get(job_id)
    if not job:
        raise HTTPException(404, f"Job {job_id} introuvable ou expire")
    return job.to_schema()


@app.get("/api/v1/download/{job_id}/{format}", tags=["Download"])
async def download(job_id: str, format: str):
    if format not in ("docx", "pdf"):
        raise HTTPException(400, "Format invalide : 'docx' ou 'pdf'")
    job = job_store.get(job_id)
    if not job:
        raise HTTPException(404, "Job introuvable ou expire")
    if job.status != "done":
        raise HTTPException(409, f"Job non termine (status: {job.status})")
    output_path = settings.TEMP_DIR / f"{job_id}_output.{format}"
    if not output_path.exists():
        raise HTTPException(404, f"Fichier {format.upper()} introuvable")
    media_types = {
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pdf":  "application/pdf",
    }
    name = job.filename.rsplit(".", 1)[0]
    return FileResponse(
        path=str(output_path),
        media_type=media_types[format],
        filename=f"{name}_corrige.{format}",
    )


@app.exception_handler(Exception)
async def global_exception_handler(request, exc):
    logger.error(f"Erreur non geree: {exc}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"detail": "Erreur interne", "vendor": "Impulse AI DocFix"},
    )
