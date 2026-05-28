"""
ai_service.py v2.0 — via n8n webhook
Impulse AI — DocFix

Chaîne :
  n8n webhook (Gemini Flash)  →  échec  →  heuristique locale
"""

import json
import re
import logging
import asyncio
import httpx
from enum import Enum
from dataclasses import dataclass
from typing import Optional

from docx import Document
from docx.shared import Pt

from .config import settings

logger = logging.getLogger("docfix.ai")


# ── Types ─────────────────────────────────────────────────────────────────────
class AIProvider(str, Enum):
    N8N       = "n8n_gemini"
    HEURISTIC = "heuristic"


class AIErrorType(str, Enum):
    QUOTA_EXCEEDED = "quota_exceeded"
    INVALID_KEY    = "invalid_key"
    UNAVAILABLE    = "unavailable"
    TIMEOUT        = "timeout"
    NONE           = "none"


@dataclass
class AIResult:
    data:         dict
    provider:     AIProvider
    error_type:   AIErrorType
    error_detail: Optional[str] = None


# ── Extraction des paragraphes ────────────────────────────────────────────────
def _extract_paragraphs(doc: Document) -> list:
    return [
        {"idx": i, "text": p.text.strip()}
        for i, p in enumerate(doc.paragraphs)
        if p.text.strip()
    ][:50]


def _build_prompt(paragraphs: list) -> str:
    return f"""Analyse ce document Word et identifie les titres et sous-titres.

Voici les premiers paragraphes (index: texte) :
{chr(10).join(f"[{p['idx']}] {p['text'][:120]}" for p in paragraphs)}

Réponds UNIQUEMENT en JSON valide avec cette structure exacte :
{{
  "doc_type": "rapport|lettre|contrat|article|autre",
  "likely_titles": [
    {{"idx": 0, "level": 1, "reason": "court, sans ponctuation finale"}},
    {{"idx": 5, "level": 2, "reason": "numéroté, style sous-section"}}
  ],
  "language": "fr|en|autre"
}}

Critères d'un titre : texte court (<80 car), pas de point final, logique de section."""


def _parse_response(raw: str) -> dict:
    clean = re.sub(r"```json|```", "", raw).strip()
    match = re.search(r"\{.*\}", clean, re.DOTALL)
    if match:
        return json.loads(match.group())
    return json.loads(clean)


def _detect_error_type(msg: str) -> AIErrorType:
    m = msg.lower()
    if "429" in m or "quota" in m or "rate" in m:
        return AIErrorType.QUOTA_EXCEEDED
    if "401" in m or "403" in m or "key" in m:
        return AIErrorType.INVALID_KEY
    if "timeout" in m:
        return AIErrorType.TIMEOUT
    return AIErrorType.UNAVAILABLE


# ── Heuristique locale ────────────────────────────────────────────────────────
def _heuristic_analysis(paragraphs: list) -> dict:
    titles = []
    for p in paragraphs:
        text = p["text"]
        if not text or len(text) > 100:
            continue
        no_dot  = not text.endswith(".")
        short   = len(text) < 80
        num     = bool(re.match(r"^\d+[\.\)]\s+\w", text))
        caps    = text == text.upper() and len(text) > 3
        roman   = bool(re.match(r"^[IVX]+[\.\)]\s+\w", text))
        if short and no_dot and (num or caps or roman):
            level = 1 if caps else 2
            titles.append({"idx": p["idx"], "level": level, "reason": "heuristique"})
    return {"doc_type": "unknown", "likely_titles": titles, "language": "fr"}


# ── Appel n8n webhook ─────────────────────────────────────────────────────────
async def _call_n8n(prompt: str) -> dict:
    """
    Envoie le prompt au webhook n8n qui appelle Gemini.
    n8n retourne directement le JSON parsé.
    """
    if not settings.N8N_GEMINI_WEBHOOK_URL:
        raise ValueError("N8N_GEMINI_WEBHOOK_URL non configurée")

    headers = {"Content-Type": "application/json"}

    # Ajouter le secret si configuré
    if settings.N8N_WEBHOOK_SECRET:
        headers["X-Webhook-Secret"] = settings.N8N_WEBHOOK_SECRET

    payload = {"prompt": prompt}

    async with httpx.AsyncClient(timeout=30.0) as client:
        response = await client.post(
            settings.N8N_GEMINI_WEBHOOK_URL,
            headers=headers,
            json=payload,
        )
        response.raise_for_status()
        data = response.json()

        # n8n peut retourner une liste ou un objet
        if isinstance(data, list):
            data = data[0]

        return data


# ── Point d'entrée principal ──────────────────────────────────────────────────
async def analyze_document(doc: Document, job_id: str = "") -> AIResult:
    """
    Analyse le document :
    1. Essaie le webhook n8n (Gemini)
    2. Fallback sur l'heuristique locale
    """
    paragraphs = _extract_paragraphs(doc)
    prompt     = _build_prompt(paragraphs)
    prefix     = f"[{job_id}] " if job_id else ""

    # ── Tentative n8n ─────────────────────────────────────────────
    if settings.N8N_GEMINI_WEBHOOK_URL:
        try:
            data = await _call_n8n(prompt)
            count = len(data.get("likely_titles", []))
            logger.info(f"{prefix}✅ n8n/Gemini OK — {count} titres")
            return AIResult(
                data=data,
                provider=AIProvider.N8N,
                error_type=AIErrorType.NONE,
            )
        except Exception as e:
            error_type   = _detect_error_type(str(e))
            error_detail = str(e)
            if error_type == AIErrorType.QUOTA_EXCEEDED:
                logger.warning(f"{prefix}⚠️ n8n/Gemini quota dépassé — fallback heuristique")
            else:
                logger.warning(f"{prefix}⚠️ n8n/Gemini indisponible ({error_type}): {e} — fallback")
    else:
        logger.info(f"{prefix}ℹ️ N8N_GEMINI_WEBHOOK_URL non configurée — heuristique locale")
        error_type   = AIErrorType.UNAVAILABLE
        error_detail = "N8N_GEMINI_WEBHOOK_URL non configurée"

    # ── Fallback heuristique ──────────────────────────────────────
    data  = _heuristic_analysis(paragraphs)
    count = len(data.get("likely_titles", []))
    logger.info(f"{prefix}🔧 Heuristique locale — {count} titres")
    return AIResult(
        data=data,
        provider=AIProvider.HEURISTIC,
        error_type=error_type,
        error_detail=error_detail if 'error_detail' in dir() else None,
    )


# ── Status checks ─────────────────────────────────────────────────────────────
async def check_gemini_status() -> dict:
    """Vérifie le webhook n8n/Gemini avec un prompt minimal."""
    if not settings.N8N_GEMINI_WEBHOOK_URL:
        return {"status": "not_configured", "provider": "n8n_gemini"}
    try:
        test_prompt = '[{"idx": 0, "text": "Test"}]\nRéponds: {"doc_type":"test","likely_titles":[],"language":"fr"}'
        data = await _call_n8n(test_prompt)
        return {
            "status":   "ok",
            "provider": "n8n_gemini",
            "webhook":  settings.N8N_GEMINI_WEBHOOK_URL,
        }
    except Exception as e:
        return {
            "status":   _detect_error_type(str(e)).value,
            "provider": "n8n_gemini",
            "detail":   str(e)[:200],
        }


async def check_fallback_status() -> dict:
    return {"status": "disabled", "provider": "none"}
