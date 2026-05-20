"""
processor.py — Moteur de traitement DOCX
Développé par Impulse AI
Chaîne complète : nettoyage → IA → styles → TOC → images → pagination → export PDF
"""

import asyncio
import logging
import subprocess
import time
from pathlib import Path
from typing import Optional

import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import lxml.etree as etree

from .config import settings
from .jobs import job_store

logger = logging.getLogger("docfix.processor")

# ── Configuration Gemini ──────────────────────────────────────────────────────
if settings.GEMINI_API_KEY:
    genai.configure(api_key=settings.GEMINI_API_KEY)


# ── Point d'entrée principal ──────────────────────────────────────────────────
async def process_document(job_id: str, input_path: Path):
    """
    Orchestre toutes les étapes de traitement.
    Appelé en arrière-plan par FastAPI BackgroundTasks.
    """
    job = job_store.get(job_id)
    if not job:
        logger.error(f"Job {job_id} introuvable")
        return

    output_docx = settings.TEMP_DIR / f"{job_id}_output.docx"
    output_pdf  = settings.TEMP_DIR / f"{job_id}_output.pdf"

    try:
        job.status = "processing"
        start_time = time.time()

        # ─ Étape 1 : Upload / réception (déjà fait, on marque done)
        _step_done(job, "upload", 0)

        # ─ Étape 2 : Parsing
        _step_start(job, "parse", "Analyse du document DOCX…")
        doc = Document(str(input_path))
        await asyncio.sleep(0)   # yield pour ne pas bloquer la boucle event
        _step_done(job, "parse", 8)

        # ─ Étape 3 : Analyse IA Gemini
        _step_start(job, "ai", "Analyse IA du contenu…")
        ai_hints = await _analyze_with_gemini(doc)
        _step_done(job, "ai", 18)

        # ─ Étape 4 : Nettoyage espaces
        _step_start(job, "clean", "Nettoyage des espaces et caractères parasites…")
        await asyncio.to_thread(_clean_whitespace, doc)
        _step_done(job, "clean", 28)

        # ─ Étape 5 : Harmonisation polices
        _step_start(job, "fonts", "Harmonisation des polices…")
        await asyncio.to_thread(_harmonize_fonts, doc)
        _step_done(job, "fonts", 38)

        # ─ Étape 6 : Détection et marquage des titres
        _step_start(job, "headings", "Détection des titres…")
        headings_count = await asyncio.to_thread(_detect_headings, doc, ai_hints)
        _step_done(job, "headings", 48)

        # ─ Étape 7 : Application des styles
        _step_start(job, "styles", "Application des styles Heading…")
        await asyncio.to_thread(_apply_styles, doc)
        _step_done(job, "styles", 56)

        # ─ Étape 8 : Table des matières
        _step_start(job, "toc", "Génération de la table des matières…")
        await asyncio.to_thread(_add_table_of_contents, doc)
        _step_done(job, "toc", 64)

        # ─ Étape 9 : Images
        _step_start(job, "images", "Redimensionnement et centrage des images…")
        images_fixed = await asyncio.to_thread(_fix_images, doc)
        _step_done(job, "images", 72)

        # ─ Étape 10 : Pagination et pieds de page
        _step_start(job, "pagination", "Ajout de la pagination…")
        await asyncio.to_thread(_add_pagination, doc)
        _step_done(job, "pagination", 80)

        # ─ Étape 11 : Export DOCX
        _step_start(job, "export_docx", "Sauvegarde du DOCX corrigé…")
        doc.save(str(output_docx))
        _step_done(job, "export_docx", 88)

        # ─ Étape 12 : Conversion PDF via LibreOffice
        _step_start(job, "export_pdf", "Conversion PDF via LibreOffice…")
        await asyncio.to_thread(_convert_to_pdf, output_docx, output_pdf)
        _step_done(job, "export_pdf", 100)

        # ─ Finalisation
        duration = round(time.time() - start_time, 1)
        job.stats = {
            "pagesCount": _count_pages(doc),
            "headingsDetected": headings_count,
            "imagesFixed": images_fixed,
            "fontsHarmonized": 1,
            "durationSeconds": duration,
        }
        job.status   = "done"
        job.progress = 100
        job.current_step = "Traitement terminé"
        logger.info(f"Job {job_id} terminé en {duration}s")

    except Exception as exc:
        logger.error(f"Erreur job {job_id}: {exc}", exc_info=True)
        job.status    = "error"
        job.error     = str(exc)
        job.current_step = "Erreur de traitement"
        # Marquer l'étape en cours comme erreur
        for step in job.steps:
            if step.status == "running":
                step.status = "error"
    finally:
        # Supprimer le fichier d'entrée
        if input_path.exists():
            input_path.unlink(missing_ok=True)


# ── Helpers de progression ─────────────────────────────────────────────────────
def _step_start(job, step_id: str, message: str):
    job.set_step_running(step_id)
    job.current_step = message
    logger.info(f"[{job.job_id}] → {message}")


def _step_done(job, step_id: str, progress: int):
    job.set_step_done(step_id)
    job.recalculate_progress()
    if progress > job.progress:
        job.progress = progress


# ── Étape 3 : Analyse Gemini ───────────────────────────────────────────────────
async def _analyze_with_gemini(doc: Document) -> dict:
    """
    Envoie un extrait du document à Gemini Flash pour identifier :
    - Les lignes qui semblent être des titres
    - Le style général du document
    """
    if not settings.GEMINI_API_KEY:
        logger.warning("GEMINI_API_KEY non configurée — analyse IA ignorée")
        return {"likely_titles": [], "doc_type": "unknown"}

    # Extraire les 50 premiers paragraphes non vides
    paragraphs = [
        {"idx": i, "text": p.text.strip(), "len": len(p.text.strip())}
        for i, p in enumerate(doc.paragraphs)
        if p.text.strip()
    ][:50]

    prompt = f"""Analyse ce document Word et identifie les titres et sous-titres.

Voici les premiers paragraphes (index: texte) :
{chr(10).join(f"[{p['idx']}] {p['text'][:120]}" for p in paragraphs)}

Réponds UNIQUEMENT en JSON valide avec cette structure :
{{
  "doc_type": "rapport|lettre|contrat|article|autre",
  "likely_titles": [
    {{"idx": 0, "level": 1, "reason": "court, sans ponctuation finale"}},
    ...
  ],
  "main_font": "nom de police détectée ou null",
  "language": "fr|en|autre"
}}

Critères d'un titre : texte court (<80 car), pas de point final, souvent en majuscules ou première lettre maj, logique de section."""

    try:
        model = genai.GenerativeModel(settings.GEMINI_MODEL)
        response = await asyncio.to_thread(
            model.generate_content,
            prompt,
            generation_config={"temperature": 0.1, "max_output_tokens": 1024},
        )
        import json, re
        raw = response.text.strip()
        # Nettoyer les balises markdown éventuelles
        raw = re.sub(r"```json|```", "", raw).strip()
        result = json.loads(raw)
        logger.info(f"Gemini : {len(result.get('likely_titles', []))} titres détectés, type={result.get('doc_type')}")
        return result
    except Exception as e:
        logger.warning(f"Gemini indisponible : {e} — fallback heuristique")
        return {"likely_titles": [], "doc_type": "unknown"}


# ── Étape 4 : Nettoyage des espaces ───────────────────────────────────────────
def _clean_whitespace(doc: Document):
    """
    - Supprime les espaces doubles
    - Supprime les sauts de ligne multiples consécutifs
    - Nettoie les espaces en début/fin de paragraphe
    """
    prev_empty = False
    paragraphs_to_delete = []

    for i, para in enumerate(doc.paragraphs):
        # Nettoyer chaque run
        for run in para.runs:
            import re
            run.text = re.sub(r' {2,}', ' ', run.text)  # espaces doubles
            run.text = re.sub(r'\t', ' ', run.text)       # tabulations → espace

        # Détecter les lignes vides consécutives
        text = para.text.strip()
        if not text:
            if prev_empty:
                paragraphs_to_delete.append(para)
            prev_empty = True
        else:
            prev_empty = False

    # Supprimer les paragraphes vides en surplus
    for para in paragraphs_to_delete:
        p = para._element
        p.getparent().remove(p)


# ── Étape 5 : Harmonisation des polices ───────────────────────────────────────
def _harmonize_fonts(doc: Document):
    """
    Unifie la police du corps de texte sur Calibri 11pt.
    Les titres (Heading) ne sont pas modifiés ici.
    """
    TARGET_FONT = "Calibri"
    TARGET_SIZE = Pt(11)

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else ""
        # Ne pas toucher aux styles de titre
        if style_name.startswith("Heading") or style_name.startswith("Title"):
            continue
        for run in para.runs:
            if run.font.size and run.font.size > Pt(14):
                continue  # probablement un titre inline, ne pas modifier
            run.font.name = TARGET_FONT
            if not run.font.size or run.font.size < Pt(8) or run.font.size > Pt(14):
                run.font.size = TARGET_SIZE


# ── Étape 6 : Détection des titres ────────────────────────────────────────────
def _detect_headings(doc: Document, ai_hints: dict) -> int:
    """
    Marque les paragraphes comme titres en combinant :
    1. Les suggestions de Gemini
    2. Une heuristique locale (taille police, texte court, etc.)
    """
    import re

    # Construire un set des index suggérés par Gemini
    ai_title_map = {}
    for t in ai_hints.get("likely_titles", []):
        ai_title_map[t["idx"]] = t.get("level", 1)

    count = 0
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text or len(text) > 150:
            continue

        # Déjà un style de titre → skip
        if para.style and para.style.name.startswith("Heading"):
            count += 1
            continue

        level = None

        # 1. Suggestion IA
        if i in ai_title_map:
            level = ai_title_map[i]

        # 2. Heuristique locale
        else:
            is_short = len(text) < 80
            no_period = not text.endswith(".")
            has_big_font = any(
                r.font.size and r.font.size >= Pt(14)
                for r in para.runs if r.font.size
            )
            is_bold_only = all(
                r.bold for r in para.runs if r.text.strip()
            ) and para.runs
            is_all_caps = text == text.upper() and len(text) > 3
            is_numbered = bool(re.match(r"^\d+[\.\)]\s+\w", text))
            roman_num   = bool(re.match(r"^[IVX]+[\.\)]\s+\w", text))

            if has_big_font and is_short and no_period:
                level = 1
            elif is_bold_only and is_short and no_period:
                level = 2
            elif (is_numbered or roman_num) and is_short:
                level = 2
            elif is_all_caps and is_short and no_period:
                level = 1

        if level:
            para._heading_level = level  # on stocke pour l'étape suivante
            count += 1

    logger.info(f"Titres détectés : {count}")
    return count


# ── Étape 7 : Application des styles Heading ──────────────────────────────────
def _apply_styles(doc: Document):
    """
    Applique les styles Heading 1/2/3 aux paragraphes marqués.
    Personnalise les styles si besoin.
    """
    for para in doc.paragraphs:
        level = getattr(para, "_heading_level", None)
        if level:
            style_name = f"Heading {min(level, 3)}"
            try:
                para.style = doc.styles[style_name]
            except KeyError:
                logger.warning(f"Style '{style_name}' absent — ignoré")


# ── Étape 8 : Table des matières ──────────────────────────────────────────────
def _add_table_of_contents(doc: Document):
    """
    Insère un champ TOC Word au début du document.
    La table se génère lors de l'ouverture dans Word / LibreOffice.
    """
    # Vérifier s'il y a des titres
    has_headings = any(
        p.style and p.style.name.startswith("Heading")
        for p in doc.paragraphs
    )
    if not has_headings:
        logger.info("Aucun titre trouvé — table des matières ignorée")
        return

    # Trouver la position d'insertion (après le premier paragraphe non vide)
    insert_before = None
    for para in doc.paragraphs:
        if para.text.strip():
            insert_before = para
            break

    # Créer le paragraphe titre "Table des matières"
    toc_title_para = OxmlElement("w:p")
    toc_title_run  = OxmlElement("w:r")
    toc_title_text = OxmlElement("w:t")
    toc_title_text.text = "Table des matières"
    toc_title_rpr = OxmlElement("w:rPr")
    bold_elem = OxmlElement("w:b")
    toc_title_rpr.append(bold_elem)
    toc_title_run.append(toc_title_rpr)
    toc_title_run.append(toc_title_text)
    toc_title_para.append(toc_title_run)

    # Créer le champ TOC
    toc_para = OxmlElement("w:p")
    toc_run  = OxmlElement("w:r")
    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = ' TOC \\o "1-3" \\h \\z \\u '
    fld_char_sep = OxmlElement("w:fldChar")
    fld_char_sep.set(qn("w:fldCharType"), "separate")
    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")

    toc_run.append(fld_char_begin)
    toc_run.append(instr_text)
    toc_run.append(fld_char_sep)
    toc_run.append(fld_char_end)
    toc_para.append(toc_run)

    # Paragraphe de séparation
    sep_para = OxmlElement("w:p")

    # Insérer avant le premier paragraphe
    if insert_before is not None:
        ref = insert_before._element
        parent = ref.getparent()
        idx = list(parent).index(ref)
        parent.insert(idx, sep_para)
        parent.insert(idx, toc_para)
        parent.insert(idx, toc_title_para)
    else:
        doc.element.body.append(toc_title_para)
        doc.element.body.append(toc_para)
        doc.element.body.append(sep_para)

    logger.info("Table des matières insérée")


# ── Étape 9 : Images ──────────────────────────────────────────────────────────
def _fix_images(doc: Document) -> int:
    """
    Pour chaque image :
    - Centre le paragraphe contenant l'image
    - Redimensionne si trop large (max 14cm de large)
    """
    MAX_WIDTH = Cm(14)
    count = 0

    for para in doc.paragraphs:
        has_image = False
        for run in para.runs:
            # Détecter les images inline
            for drawing in run._element.findall(
                ".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline"
            ) + run._element.findall(
                ".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor"
            ):
                has_image = True
                # Chercher l'extent (dimensions)
                extent = drawing.find(
                    "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
                )
                if extent is not None:
                    cx = int(extent.get("cx", 0))
                    # 914400 EMUs = 1 pouce = 2.54 cm
                    max_cx = int(MAX_WIDTH.emu)
                    if cx > max_cx:
                        cy = int(extent.get("cy", 0))
                        ratio = max_cx / cx
                        extent.set("cx", str(max_cx))
                        extent.set("cy", str(int(cy * ratio)))
                        count += 1

        if has_image:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    logger.info(f"Images traitées : {count} redimensionnées")
    return count


# ── Étape 10 : Pagination et pied de page ────────────────────────────────────
def _add_pagination(doc: Document):
    """
    Ajoute un pied de page avec :
    - Nom du document (centré)
    - Numéro de page (droite)
    """
    from docx.oxml.ns import nsmap

    # S'assurer qu'il y a une section
    if not doc.sections:
        return

    section = doc.sections[-1]
    footer = section.footer

    # Vider le footer existant
    for para in footer.paragraphs:
        for run in para.runs:
            run.text = ""

    # Créer le paragraphe du pied de page
    if footer.paragraphs:
        footer_para = footer.paragraphs[0]
    else:
        footer_para = footer.add_paragraph()

    footer_para.clear()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Ajouter "Impulse AI | DocFix  —  Page X / Y"
    run1 = footer_para.add_run("Impulse AI | DocFix  —  Page ")
    run1.font.size = Pt(8)
    run1.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    # Champ PAGE
    _add_field(footer_para, "PAGE")

    run2 = footer_para.add_run(" / ")
    run2.font.size = Pt(8)
    run2.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    # Champ NUMPAGES
    _add_field(footer_para, "NUMPAGES")

    # Activer le pied de page
    sectPr = section._sectPr
    pg_sz = sectPr.find(qn("w:pgSz"))
    footer_ref = OxmlElement("w:footerReference")
    footer_ref.set(qn("w:type"), "default")


def _add_field(para, field_name: str):
    """Insère un champ Word (PAGE, NUMPAGES…) dans un paragraphe."""
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "16")  # 8pt = 16 half-points
    rpr.append(sz)
    run.append(rpr)

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run.append(fld_begin)
    para._p.append(run)

    run2 = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_name} "
    run2.append(instr)
    para._p.append(run2)

    run3 = OxmlElement("w:r")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run3.append(fld_end)
    para._p.append(run3)


# ── Étape 12 : Conversion PDF via LibreOffice ─────────────────────────────────
def _convert_to_pdf(docx_path: Path, pdf_path: Path):
    """
    Convertit le DOCX en PDF via LibreOffice headless.
    Requis sur Render : LibreOffice installé via apt.
    """
    cmd = [
        settings.LIBREOFFICE_PATH,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(pdf_path.parent),
        str(docx_path),
    ]
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice error: {result.stderr}")

        # LibreOffice nomme le fichier en fonction du docx
        generated = pdf_path.parent / (docx_path.stem + ".pdf")
        if generated.exists() and generated != pdf_path:
            generated.rename(pdf_path)

        if not pdf_path.exists():
            raise FileNotFoundError("Le fichier PDF n'a pas été généré")

        logger.info(f"PDF généré : {pdf_path} ({pdf_path.stat().st_size // 1024} Ko)")

    except subprocess.TimeoutExpired:
        raise RuntimeError("LibreOffice : timeout dépassé (120s)")
    except FileNotFoundError:
        raise RuntimeError(
            f"LibreOffice introuvable à '{settings.LIBREOFFICE_PATH}'. "
            "Vérifier l'installation sur Render."
        )


# ── Utilitaires ───────────────────────────────────────────────────────────────
def _count_pages(doc: Document) -> int:
    """Estimation du nombre de pages (approximatif sans rendu)."""
    total_lines = sum(
        max(1, len(p.text) // 80)
        for p in doc.paragraphs
        if p.text.strip()
    )
    return max(1, total_lines // 40)
