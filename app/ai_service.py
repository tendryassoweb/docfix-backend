"""
ai_service.py — Couche IA avec fallback générique et détecteur d'erreurs
Impulse AI — DocFix

Chaîne :
  Gemini Flash  →  (quota/erreur)  →  Fallback API  →  (erreur)  →  Heuristique locale
"""

import json
import re
import logging
import asyncio
import httpx
from enum import Enum
from dataclasses import dataclass
from typing import Optional

import google.generativeai as genai
from docx import Document

from .config import settings

logger = logging.getLogger("docfix.ai")


# ── Types ─────────────────────────────────────────────────────────────────────

class AIProvider(str, Enum):
    GEMINI    = "gemini"
    FALLBACK  = "fallback"
    HEURISTIC = "heuristic"  # pas d'IA, règles locales uniquement


class AIErrorType(str, Enum):
    QUOTA_EXCEEDED  = "quota_exceeded"
    INVALID_KEY     = "invalid_key"
    UNAVAILABLE     = "unavailable"
    TIMEOUT         = "timeout"
    PARSE_ERROR     = "parse_error"
    NONE            = "none"


@dataclass
class AIResult:
    """Résultat de l'analyse IA avec métadonnées de diagnostic."""
    data: dict                          # résultat JSON de l'analyse
    provider: AIProvider                # qui a répondu
    error_type: AIErrorType             # erreur rencontrée (si fallback activé)
    error_detail: Optional[str] = None  # message d'erreur brut
    tokens_used: Optional[int]  = None  # tokens consommés si disponible


# ── Configuration Gemini ──────────────────────────────────────────────────────
if settings.GEMINI_API_KEY:
    genai.configure(api_key=settings.GEMINI_API_KEY)


# ── Prompt partagé ────────────────────────────────────────────────────────────
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


def _parse_ai_response(raw: str) -> dict:
    """Nettoie et parse la réponse JSON de n'importe quelle IA."""
    # Supprimer les balises markdown
    clean = re.sub(r"```json|```", "", raw).strip()
    # Extraire le premier objet JSON valide
    match = re.search(r"\{.*\}", clean, re.DOTALL)
    if match:
        return json.loads(match.group())
    return json.loads(clean)


def _extract_paragraphs(doc: Document) -> list:
    """Extrait les 50 premiers paragraphes non vides pour l'analyse."""
    return [
        {"idx": i, "text": p.text.strip()}
        for i, p in enumerate(doc.paragraphs)
        if p.text.strip()
    ][:50]


def _detect_error_type(error_msg: str) -> AIErrorType:
    """Classifie le type d'erreur à partir du message."""
    msg = error_msg.lower()
    if "429" in msg or "quota" in msg or "rate limit" in msg or "resource_exhausted" in msg:
        return AIErrorType.QUOTA_EXCEEDED
    if "401" in msg or "403" in msg or "api key" in msg or "invalid" in msg and "key" in msg:
        return AIErrorType.INVALID_KEY
    if "timeout" in msg or "timed out" in msg:
        return AIErrorType.TIMEOUT
    if "503" in msg or "unavailable" in msg or "overloaded" in msg:
        return AIErrorType.UNAVAILABLE
    return AIErrorType.UNAVAILABLE


# ── Provider 1 : Gemini ───────────────────────────────────────────────────────
async def _call_gemini(prompt: str) -> dict:
    """Appel à Gemini Flash. Lève une exception typée si erreur."""
    model = genai.GenerativeModel(settings.GEMINI_MODEL)
    response = await asyncio.to_thread(
        model.generate_content,
        prompt,
        generation_config={"temperature": 0.1, "max_output_tokens": 1024},
    )
    return _parse_ai_response(response.text)


# ── Provider 2 : Fallback générique (format OpenAI compatible) ────────────────
async def _call_fallback(prompt: str) -> dict:
    """
    Appel à l'API fallback configurée.
    Compatible avec toute API au format OpenAI (Mistral, Groq, Together, etc.)
    """
    headers = {
        "Authorization": f"Bearer {settings.FALLBACK_AI_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": settings.FALLBACK_AI_MODEL,
        "messages": [
            {
                "role": "system",
                "content": "Tu es un assistant qui analyse des documents Word. Réponds toujours en JSON valide uniquement, sans markdown."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "temperature": 0.1,
        "max_tokens": 1024,
    }

    async with httpx.AsyncClient(timeout=30.0) as client:
        response = await client.post(
            f"{settings.FALLBACK_AI_BASE_URL}/chat/completions",
            headers=headers,
            json=payload,
        )
        response.raise_for_status()
        data = response.json()
        raw = data["choices"][0]["message"]["content"]
        return _parse_ai_response(raw)


# ── Provider 3 : Heuristique locale (aucune IA) ───────────────────────────────
def _heuristic_analysis(paragraphs: list) -> dict:
    """
    Détection de titres par règles locales.
    Utilisé quand toutes les IA sont indisponibles.
    """
    likely_titles = []
    for p in paragraphs:
        text = p["text"]
        if not text or len(text) > 100:
            continue
        no_period   = not text.endswith(".")
        is_short    = len(text) < 80
        is_numbered = bool(re.match(r"^\d+[\.\)]\s+\w", text))
        is_caps     = text == text.upper() and len(text) > 3
        is_roman    = bool(re.match(r"^[IVX]+[\.\)]\s+\w", text))

        if is_short and no_period and (is_numbered or is_caps or is_roman):
            level = 1 if is_caps else 2
            likely_titles.append({
                "idx": p["idx"],
                "level": level,
                "reason": "heuristique locale"
            })

    return {
        "doc_type": "unknown",
        "likely_titles": likely_titles,
        "language": "fr",
    }


# ── Point d'entrée principal ──────────────────────────────────────────────────
async def analyze_document(doc: Document, job_id: str = "") -> AIResult:
    """
    Analyse le document avec la chaîne de fallback :
    Gemini → Fallback API → Heuristique locale

    Retourne toujours un AIResult valide, jamais d'exception.
    """
    paragraphs = _extract_paragraphs(doc)
    prompt     = _build_prompt(paragraphs)
    prefix     = f"[{job_id}] " if job_id else ""

    # ── Tentative 1 : Gemini ──────────────────────────────────────
    if settings.GEMINI_API_KEY:
        try:
            data = await _call_gemini(prompt)
            titles_count = len(data.get("likely_titles", []))
            logger.info(f"{prefix}✅ Gemini OK — {titles_count} titres détectés")
            return AIResult(
                data=data,
                provider=AIProvider.GEMINI,
                error_type=AIErrorType.NONE,
            )
        except Exception as e:
            error_type   = _detect_error_type(str(e))
            error_detail = str(e)

            # Log selon le type d'erreur
            if error_type == AIErrorType.QUOTA_EXCEEDED:
                logger.warning(f"{prefix}⚠️  Gemini quota dépassé — activation du fallback")
            elif error_type == AIErrorType.INVALID_KEY:
                logger.error(f"{prefix}❌ Gemini clé API invalide — vérifier GEMINI_API_KEY")
            else:
                logger.warning(f"{prefix}⚠️  Gemini indisponible ({error_type}) — activation du fallback")
    else:
        error_type   = AIErrorType.INVALID_KEY
        error_detail = "GEMINI_API_KEY non configurée"
        logger.warning(f"{prefix}⚠️  Gemini non configuré")

    # ── Tentative 2 : Fallback API ────────────────────────────────
    fallback_ready = (
        settings.FALLBACK_AI_ENABLED
        and settings.FALLBACK_AI_API_KEY
        and settings.FALLBACK_AI_BASE_URL
        and settings.FALLBACK_AI_MODEL
    )

    if fallback_ready:
        try:
            data = await _call_fallback(prompt)
            titles_count = len(data.get("likely_titles", []))
            logger.info(f"{prefix}✅ {settings.FALLBACK_AI_NAME} OK — {titles_count} titres détectés")
            return AIResult(
                data=data,
                provider=AIProvider.FALLBACK,
                error_type=error_type,        # on garde l'erreur Gemini pour info
                error_detail=error_detail,
            )
        except Exception as e2:
            fallback_error = _detect_error_type(str(e2))
            logger.warning(
                f"{prefix}⚠️  {settings.FALLBACK_AI_NAME} indisponible ({fallback_error}) "
                f"— activation de l'heuristique locale"
            )
    else:
        if settings.FALLBACK_AI_ENABLED:
            logger.warning(f"{prefix}⚠️  Fallback activé mais mal configuré — vérifier les variables")
        else:
            logger.info(f"{prefix}ℹ️  Fallback désactivé — heuristique locale utilisée")

    # ── Tentative 3 : Heuristique locale ─────────────────────────
    data = _heuristic_analysis(paragraphs)
    titles_count = len(data.get("likely_titles", []))
    logger.info(f"{prefix}🔧 Heuristique locale — {titles_count} titres détectés")
    return AIResult(
        data=data,
        provider=AIProvider.HEURISTIC,
        error_type=error_type,
        error_detail=error_detail,
    )


# ── Test de connectivité ───────────────────────────────────────────────────────
async def check_gemini_status() -> dict:
    """Vérifie si Gemini est accessible et le quota disponible."""
    if not settings.GEMINI_API_KEY:
        return {"status": "not_configured", "provider": "gemini"}
    try:
        model = genai.GenerativeModel(settings.GEMINI_MODEL)
        await asyncio.to_thread(
            model.generate_content,
            "Réponds juste: OK",
            generation_config={"max_output_tokens": 5},
        )
        return {"status": "ok", "provider": "gemini", "model": settings.GEMINI_MODEL}
    except Exception as e:
        error_type = _detect_error_type(str(e))
        return {
            "status": error_type.value,
            "provider": "gemini",
            "model": settings.GEMINI_MODEL,
            "detail": str(e)[:200],
        }


async def check_fallback_status() -> dict:
    """Vérifie si le fallback est accessible."""
    if not settings.FALLBACK_AI_ENABLED:
        return {"status": "disabled", "provider": settings.FALLBACK_AI_NAME}
    if not all([settings.FALLBACK_AI_API_KEY, settings.FALLBACK_AI_BASE_URL, settings.FALLBACK_AI_MODEL]):
        return {"status": "misconfigured", "provider": settings.FALLBACK_AI_NAME}
    try:
        headers = {
            "Authorization": f"Bearer {settings.FALLBACK_AI_API_KEY}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": settings.FALLBACK_AI_MODEL,
            "messages": [{"role": "user", "content": "Réponds juste: OK"}],
            "max_tokens": 5,
        }
        async with httpx.AsyncClient(timeout=10.0) as client:
            response = await client.post(
                f"{settings.FALLBACK_AI_BASE_URL}/chat/completions",
                headers=headers,
                json=payload,
            )
            response.raise_for_status()
        return {
            "status": "ok",
            "provider": settings.FALLBACK_AI_NAME,
            "model": settings.FALLBACK_AI_MODEL,
        }
    except Exception as e:
        error_type = _detect_error_type(str(e))
        return {
            "status": error_type.value,
            "provider": settings.FALLBACK_AI_NAME,
            "detail": str(e)[:200],
        }
