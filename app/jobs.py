"""
jobs.py — Gestion des jobs en mémoire (thread-safe)
Pour la V1 : stockage en mémoire. Pour la V2 : migrer vers Redis.
"""

import time
import threading
import logging
from typing import Optional
from dataclasses import dataclass, field
from .config import settings

logger = logging.getLogger("docfix.jobs")

# ── Étapes de traitement ──────────────────────────────────────────────────────
STEPS_DEFINITION = [
    ("upload",      "Réception du fichier"),
    ("parse",       "Analyse du document"),
    ("ai",          "Analyse IA (Gemini)"),
    ("clean",       "Nettoyage et espaces"),
    ("fonts",       "Harmonisation des polices"),
    ("headings",    "Détection des titres"),
    ("styles",      "Application des styles"),
    ("toc",         "Génération table des matières"),
    ("images",      "Redimensionnement images"),
    ("pagination",  "Pagination & pieds de page"),
    ("export_docx", "Export DOCX final"),
    ("export_pdf",  "Conversion PDF"),
]


@dataclass
class Step:
    id: str
    label: str
    status: str = "pending"  # pending | running | done | error

    def to_dict(self) -> dict:
        return {"id": self.id, "label": self.label, "status": self.status}


@dataclass
class Job:
    job_id: str
    filename: str
    status: str = "queued"    # queued | processing | done | error
    progress: int = 0
    current_step: str = ""
    error: Optional[str] = None
    created_at: float = field(default_factory=time.time)
    stats: Optional[dict] = None
    steps: list = field(default_factory=lambda: [
        Step(id=s[0], label=s[1]) for s in STEPS_DEFINITION
    ])

    def to_schema(self) -> dict:
        return {
            "jobId": self.job_id,
            "status": self.status,
            "progress": self.progress,
            "currentStep": self.current_step,
            "steps": [s.to_dict() for s in self.steps],
            "error": self.error,
            "result": {
                "jobId": self.job_id,
                "status": self.status,
                "stats": self.stats,
            } if self.status == "done" else None,
        }

    def set_step_running(self, step_id: str):
        for s in self.steps:
            if s.id == step_id:
                s.status = "running"
            elif s.status == "running":
                s.status = "pending"  # ne laisse qu'une étape active

    def set_step_done(self, step_id: str):
        for s in self.steps:
            if s.id == step_id:
                s.status = "done"

    def set_step_error(self, step_id: str):
        for s in self.steps:
            if s.id == step_id:
                s.status = "error"

    def recalculate_progress(self):
        done = sum(1 for s in self.steps if s.status == "done")
        self.progress = int(done / len(self.steps) * 100)

    def is_expired(self) -> bool:
        return (time.time() - self.created_at) > settings.JOB_EXPIRY_SECONDS


# ── Store thread-safe ─────────────────────────────────────────────────────────
class JobStore:
    def __init__(self):
        self._jobs: dict[str, Job] = {}
        self._lock = threading.Lock()

    def create(self, job_id: str, filename: str) -> Job:
        job = Job(job_id=job_id, filename=filename)
        with self._lock:
            self._jobs[job_id] = job
        return job

    def get(self, job_id: str) -> Optional[Job]:
        with self._lock:
            job = self._jobs.get(job_id)
            if job and job.is_expired():
                self._cleanup_job(job)
                del self._jobs[job_id]
                return None
            return job

    def update(self, job_id: str, **kwargs):
        with self._lock:
            job = self._jobs.get(job_id)
            if job:
                for k, v in kwargs.items():
                    setattr(job, k, v)

    def cleanup_expired(self) -> int:
        """Supprime les jobs expirés. Retourne le nombre supprimé."""
        with self._lock:
            expired = [jid for jid, j in self._jobs.items() if j.is_expired()]
            for jid in expired:
                self._cleanup_job(self._jobs[jid])
                del self._jobs[jid]
        if expired:
            logger.info(f"Suppression de {len(expired)} jobs expirés")
        return len(expired)

    def _cleanup_job(self, job: Job):
        """Supprime les fichiers temporaires d'un job."""
        from .config import settings
        for suffix in ["_input.docx", "_output.docx", "_output.pdf"]:
            path = settings.TEMP_DIR / f"{job.job_id}{suffix}"
            if path.exists():
                try:
                    path.unlink()
                except Exception as e:
                    logger.warning(f"Impossible de supprimer {path}: {e}")


# Instance globale
job_store = JobStore()
