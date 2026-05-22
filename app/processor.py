"""
processor.py v1.5
Impulse AI — DocFix

Fix critique : _heading_level stocké dans un dict partagé
au lieu d'attributs Python temporaires sur les objets para
(qui ne survivent pas aux appels asyncio.to_thread séparés)
"""

import asyncio, logging, subprocess, time, re, os
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.style import WD_STYLE_TYPE

from .config import settings
from .jobs import job_store
from .ai_service import analyze_document, AIProvider

logger = logging.getLogger("docfix.processor")

COLOR_H1     = RGBColor(0x1F, 0x35, 0x64)
COLOR_H2     = RGBColor(0x2E, 0x74, 0xB5)
COLOR_H3     = RGBColor(0x40, 0x40, 0x40)
COLOR_BODY   = RGBColor(0x26, 0x26, 0x26)
COLOR_FOOTER = RGBColor(0x80, 0x80, 0x80)
FONT_HEAD    = "Calibri"
FONT_BODY    = "Calibri"


async def process_document(job_id: str, input_path: Path):
    job = job_store.get(job_id)
    if not job:
        return

    output_docx = settings.TEMP_DIR / f"{job_id}_output.docx"
    output_pdf  = settings.TEMP_DIR / f"{job_id}_output.pdf"

    try:
        job.status = "processing"
        start = time.time()
        _step_done(job, "upload", 0)

        _step_start(job, "parse", "Analyse du document...")
        doc = Document(str(input_path))
        await asyncio.sleep(0)
        _step_done(job, "parse", 8)

        _step_start(job, "ai", "Analyse IA...")
        ai_result = await analyze_document(doc, job_id)
        _step_done(job, "ai", 18)

        _step_start(job, "clean", "Nettoyage des espaces...")
        await asyncio.to_thread(_clean_whitespace, doc)
        _step_done(job, "clean", 26)

        _step_start(job, "fonts", "Harmonisation des polices...")
        await asyncio.to_thread(_harmonize_fonts, doc)
        _step_done(job, "fonts", 34)

        # ── Détection → retourne un dict {para_index: level} ──────────
        _step_start(job, "headings", "Detection des titres...")
        heading_map = await asyncio.to_thread(
            _detect_headings, doc, ai_result.data
        )
        n_heads = len(heading_map)
        logger.info(f"[{job_id}] heading_map={n_heads} entrées : {dict(list(heading_map.items())[:5])}")
        _step_done(job, "headings", 44)

        # ── Styles + couleurs + TOC — tout dans UN SEUL to_thread ─────
        # On passe heading_map explicitement pour éviter
        # la perte des attributs entre threads
        _step_start(job, "styles", "Application styles, couleurs et TOC...")

        toc_ok = await asyncio.to_thread(
            _apply_all_formatting, doc, heading_map
        )
        _step_done(job, "styles", 54)

        # Marquer TOC done séparément pour le frontend
        job.set_step_done("toc")
        job.recalculate_progress()
        job.progress = max(job.progress, 63)

        _step_start(job, "images", "Redimensionnement images...")
        n_img = await asyncio.to_thread(_fix_images, doc)
        _step_done(job, "images", 71)

        _step_start(job, "pagination", "Pagination...")
        await asyncio.to_thread(_add_pagination, doc)
        _step_done(job, "pagination", 79)

        _step_start(job, "export_docx", "Sauvegarde DOCX...")
        doc.save(str(output_docx))
        _step_done(job, "export_docx", 88)

        _step_start(job, "export_pdf", "Conversion PDF...")
        await asyncio.to_thread(_to_pdf, output_docx, output_pdf)
        _step_done(job, "export_pdf", 100)

        duration = round(time.time() - start, 1)
        job.stats = {
            "pagesCount":       _count_pages(doc),
            "headingsDetected": n_heads,
            "imagesFixed":      n_img,
            "fontsHarmonized":  1,
            "durationSeconds":  duration,
            "aiProvider":       ai_result.provider.value,
            "tocInserted":      toc_ok,
        }
        job.status       = "done"
        job.progress     = 100
        job.current_step = "Traitement termine"
        logger.info(f"Job {job_id} OK en {duration}s — heads={n_heads} toc={toc_ok}")

    except Exception as exc:
        logger.error(f"Erreur job {job_id}: {exc}", exc_info=True)
        job.status       = "error"
        job.error        = str(exc)
        job.current_step = "Erreur de traitement"
        for s in job.steps:
            if s.status == "running":
                s.status = "error"
    finally:
        input_path.unlink(missing_ok=True)


# ── _apply_all_formatting : styles + TOC dans le même thread ────────────────
def _apply_all_formatting(doc: Document, heading_map: dict) -> bool:
    """
    Regroupe setup_styles + apply_heading_styles + add_toc dans
    un seul appel synchrone pour garantir que heading_map est accessible.
    Retourne True si la TOC a été insérée.
    """
    _setup_styles(doc)
    _apply_heading_styles_with_map(doc, heading_map)
    return _add_toc_with_map(doc, heading_map)


# ── Helpers progression ───────────────────────────────────────────────────────
def _step_start(job, sid, msg):
    job.set_step_running(sid)
    job.current_step = msg
    logger.info(f"[{job.job_id}] {msg}")

def _step_done(job, sid, pct):
    job.set_step_done(sid)
    job.recalculate_progress()
    if pct > job.progress:
        job.progress = pct


# ── Nettoyage ─────────────────────────────────────────────────────────────────
def _clean_whitespace(doc):
    prev_empty = False
    to_del = []
    for para in doc.paragraphs:
        for run in para.runs:
            run.text = re.sub(r' {2,}', ' ', run.text)
            run.text = re.sub(r'\t+', ' ', run.text)
        if not para.text.strip():
            if prev_empty:
                to_del.append(para)
            prev_empty = True
        else:
            prev_empty = False
    for p in to_del:
        p._element.getparent().remove(p._element)


# ── Polices ───────────────────────────────────────────────────────────────────
def _harmonize_fonts(doc):
    for para in doc.paragraphs:
        sn = para.style.name if para.style else ""
        if any(sn.startswith(x) for x in ["Heading", "Title", "TOC"]):
            continue
        for run in para.runs:
            if run.font.size and run.font.size > Pt(14):
                continue
            run.font.name      = FONT_BODY
            run.font.color.rgb = COLOR_BODY
            if not run.font.size or not (Pt(8) <= run.font.size <= Pt(14)):
                run.font.size = Pt(11)


# ── Détection titres → dict {index_para: level} ───────────────────────────────
def _detect_headings(doc: Document, ai_hints: dict) -> dict:
    """
    Retourne un dictionnaire {para_index: heading_level}.
    On utilise l'index (int) comme clé — stable entre threads.
    """
    ai_map = {t["idx"]: t.get("level", 1) for t in ai_hints.get("likely_titles", [])}
    result = {}

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text or len(text) > 150:
            continue

        sn = para.style.name if para.style else ""

        # Déjà Heading dans le doc original
        if sn.startswith("Heading"):
            try:
                lvl = int(sn.split()[-1])
            except (ValueError, IndexError):
                lvl = 1
            result[i] = lvl
            continue

        # Suggestion IA
        level = ai_map.get(i)

        # Heuristique locale
        if level is None:
            short  = len(text) < 80
            no_dot = not text.endswith(".")
            big_f  = any(r.font.size and r.font.size >= Pt(14) for r in para.runs if r.font.size)
            bold   = bool(para.runs) and all(r.bold for r in para.runs if r.text.strip())
            caps   = text == text.upper() and len(text) > 3
            num    = bool(re.match(r"^\d+[\.\)]\s+\w", text))
            roman  = bool(re.match(r"^[IVX]+[\.\)]\s+\w", text))

            if   big_f and short and no_dot:  level = 1
            elif bold  and short and no_dot:  level = 2
            elif (num  or roman) and short:   level = 2
            elif caps  and short and no_dot:  level = 1

        if level:
            result[i] = level

    logger.info(f"_detect_headings : {len(result)} titres dans {len(doc.paragraphs)} paragraphes")
    return result


# ── Setup styles Heading dans le document ────────────────────────────────────
def _setup_styles(doc):
    cfgs = {
        "Heading 1": dict(size=Pt(18), color=COLOR_H1, bold=True,  italic=False,
                          caps=True,  before=Pt(24), after=Pt(8)),
        "Heading 2": dict(size=Pt(14), color=COLOR_H2, bold=True,  italic=False,
                          caps=False, before=Pt(16), after=Pt(6)),
        "Heading 3": dict(size=Pt(12), color=COLOR_H3, bold=True,  italic=True,
                          caps=False, before=Pt(12), after=Pt(4)),
    }
    for name, c in cfgs.items():
        try:
            try:
                st = doc.styles[name]
            except KeyError:
                st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            st.font.name      = FONT_HEAD
            st.font.size      = c["size"]
            st.font.color.rgb = c["color"]
            st.font.bold      = c["bold"]
            st.font.italic    = c["italic"]
            if c["caps"]:
                st.font.all_caps = True
            pf = st.paragraph_format
            pf.space_before      = c["before"]
            pf.space_after       = c["after"]
            pf.keep_with_next    = True
            pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            pf.alignment         = WD_ALIGN_PARAGRAPH.LEFT
        except Exception as e:
            logger.warning(f"Style {name}: {e}")
    try:
        n = doc.styles["Normal"]
        n.font.name      = FONT_BODY
        n.font.size      = Pt(11)
        n.font.color.rgb = COLOR_BODY
        n.paragraph_format.space_after       = Pt(6)
        n.paragraph_format.alignment         = WD_ALIGN_PARAGRAPH.JUSTIFY
        n.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    except Exception:
        pass


# ── Appliquer styles + couleurs avec heading_map ─────────────────────────────
def _apply_heading_styles_with_map(doc: Document, heading_map: dict):
    color_map = {1: COLOR_H1,  2: COLOR_H2,  3: COLOR_H3}
    size_map  = {1: Pt(18),    2: Pt(14),    3: Pt(12)}

    for i, para in enumerate(doc.paragraphs):
        lvl = heading_map.get(i)
        if not lvl:
            continue
        lvl = min(lvl, 3)

        # Style paragraphe
        try:
            para.style = doc.styles[f"Heading {lvl}"]
        except KeyError:
            pass

        # Forcer couleur + taille sur chaque run via XML
        for run in para.runs:
            _force_run_formatting(
                run,
                color    = color_map[lvl],
                size     = size_map[lvl],
                bold     = True,
                italic   = (lvl == 3),
                all_caps = (lvl == 1),
            )

        # Espacement
        pf = para.paragraph_format
        pf.space_before   = {1: Pt(24), 2: Pt(16), 3: Pt(12)}[lvl]
        pf.space_after    = {1: Pt(8),  2: Pt(6),  3: Pt(4)}[lvl]
        pf.keep_with_next = True

        # Filet bleu sous H1
        if lvl == 1:
            _set_para_border(para, "bottom", "1F3564", "8")


def _force_run_formatting(run, color: RGBColor, size: Pt,
                           bold=True, italic=False, all_caps=False):
    """Force le formatage XML en supprimant les overrides thème Word."""
    rpr = run._r.get_or_add_rPr()

    # Police
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rpr.insert(0, rFonts)
    for attr in ["w:ascii", "w:hAnsi", "w:cs"]:
        rFonts.set(qn(attr), FONT_HEAD)

    # Taille
    half = str(int(size.pt * 2))
    for tag in ["w:sz", "w:szCs"]:
        el = rpr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag); rpr.append(el)
        el.set(qn("w:val"), half)

    # Gras
    b = rpr.find(qn("w:b"))
    if bold and b is None:
        rpr.append(OxmlElement("w:b"))
    elif not bold and b is not None:
        rpr.remove(b)

    # Italique
    i_el = rpr.find(qn("w:i"))
    if italic and i_el is None:
        rpr.append(OxmlElement("w:i"))
    elif not italic and i_el is not None:
        rpr.remove(i_el)

    # Majuscules
    caps_el = rpr.find(qn("w:caps"))
    if all_caps and caps_el is None:
        rpr.append(OxmlElement("w:caps"))
    elif not all_caps and caps_el is not None:
        rpr.remove(caps_el)

    # Couleur — supprimer themeColor et forcer RGB
    color_el = rpr.find(qn("w:color"))
    if color_el is None:
        color_el = OxmlElement("w:color"); rpr.append(color_el)
    color_el.set(qn("w:val"), f"{color.red:02X}{color.green:02X}{color.blue:02X}")
    for attr in [qn("w:themeColor"), qn("w:themeTint"), qn("w:themeShade")]:
        if attr in color_el.attrib:
            del color_el.attrib[attr]

    # Supprimer highlight
    hl = rpr.find(qn("w:highlight"))
    if hl is not None:
        rpr.remove(hl)


def _set_para_border(para, side, color_hex, sz="6"):
    pPr  = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    el   = OxmlElement(f"w:{side}")
    el.set(qn("w:val"), "single"); el.set(qn("w:sz"), sz)
    el.set(qn("w:space"), "4");    el.set(qn("w:color"), color_hex)
    pBdr.append(el)
    old = pPr.find(qn("w:pBdr"))
    if old is not None:
        pPr.remove(old)
    pPr.append(pBdr)


# ── TOC avec heading_map ──────────────────────────────────────────────────────
def _add_toc_with_map(doc: Document, heading_map: dict) -> bool:
    """Insère la TOC. Utilise heading_map au lieu de _heading_level."""
    if not heading_map:
        logger.warning(f"heading_map vide — TOC ignorée")
        return False

    logger.info(f"_add_toc : {len(heading_map)} titres trouvés dans heading_map")

    first = next((p for p in doc.paragraphs if p.text.strip()), None)

    # Titre centré
    title_p   = OxmlElement("w:p")
    title_ppr = OxmlElement("w:pPr")
    jc = OxmlElement("w:jc"); jc.set(qn("w:val"), "center")
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "480"); sp.set(qn("w:after"), "240")
    title_ppr.append(jc); title_ppr.append(sp)
    title_p.append(title_ppr)
    title_r   = OxmlElement("w:r")
    title_rpr = OxmlElement("w:rPr")
    for tag, val in [("w:b", None), ("w:caps", None),
                     ("w:sz", "32"), ("w:color", "1F3564")]:
        el = OxmlElement(tag)
        if val: el.set(qn("w:val"), val)
        title_rpr.append(el)
    title_r.append(title_rpr)
    title_t = OxmlElement("w:t"); title_t.text = "Table des matieres"
    title_r.append(title_t); title_p.append(title_r)

    # Champ TOC dirty=true
    fld_p = OxmlElement("w:p")
    r1 = OxmlElement("w:r")
    fc1 = OxmlElement("w:fldChar")
    fc1.set(qn("w:fldCharType"), "begin")
    fc1.set(qn("w:dirty"), "true")
    r1.append(fc1); fld_p.append(r1)

    r2 = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = ' TOC \\o "1-3" \\h \\z \\u '
    r2.append(it); fld_p.append(r2)

    r3 = OxmlElement("w:r")
    fc3 = OxmlElement("w:fldChar"); fc3.set(qn("w:fldCharType"), "separate")
    r3.append(fc3); fld_p.append(r3)

    r4 = OxmlElement("w:r")
    ph = OxmlElement("w:t"); ph.text = "[ Table des matieres ]"
    r4.append(ph); fld_p.append(r4)

    r5 = OxmlElement("w:r")
    fc5 = OxmlElement("w:fldChar"); fc5.set(qn("w:fldCharType"), "end")
    r5.append(fc5); fld_p.append(r5)

    # Saut de page
    pb_p = OxmlElement("w:p")
    pb_r = OxmlElement("w:r")
    pb_e = OxmlElement("w:br"); pb_e.set(qn("w:type"), "page")
    pb_r.append(pb_e); pb_p.append(pb_r)

    elems = [title_p, fld_p, pb_p]
    if first is not None:
        ref = first._element
        par = ref.getparent()
        idx = list(par).index(ref)
        for el in reversed(elems):
            par.insert(idx, el)
    else:
        for el in elems:
            doc.element.body.append(el)

    logger.info(f"TOC inseree avec {len(heading_map)} titres")
    return True


# ── Images ────────────────────────────────────────────────────────────────────
def _fix_images(doc) -> int:
    MAX = Cm(14)
    NS  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    n   = 0
    for para in doc.paragraphs:
        has_img = False
        for run in para.runs:
            for tag in [f"{{{NS}}}inline", f"{{{NS}}}anchor"]:
                for d in run._element.findall(f".//{tag}"):
                    has_img = True
                    ext = d.find(f"{{{NS}}}extent")
                    if ext is not None:
                        cx = int(ext.get("cx", 0))
                        mx = int(MAX.emu)
                        if cx > mx:
                            cy = int(ext.get("cy", 0))
                            ext.set("cx", str(mx))
                            ext.set("cy", str(int(cy * mx / cx)))
                            n += 1
        if has_img:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return n


# ── Pagination ────────────────────────────────────────────────────────────────
def _add_pagination(doc):
    if not doc.sections:
        return
    for sec in doc.sections:
        sec.top_margin    = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin   = Cm(2.5)
        sec.right_margin  = Cm(2.5)
    footer = doc.sections[-1].footer
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.clear()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_para_border(fp, "top", "CCCCCC", "4")
    r1 = fp.add_run("Impulse AI  |  DocFix          Page ")
    r1.font.name = FONT_BODY; r1.font.size = Pt(8); r1.font.color.rgb = COLOR_FOOTER
    _add_field(fp, "PAGE")
    r2 = fp.add_run(" / ")
    r2.font.name = FONT_BODY; r2.font.size = Pt(8); r2.font.color.rgb = COLOR_FOOTER
    _add_field(fp, "NUMPAGES")


def _add_field(para, name):
    for ftype, content in [("begin", None), (None, name), ("end", None)]:
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        sz  = OxmlElement("w:sz"); sz.set(qn("w:val"), "16")
        col = OxmlElement("w:color"); col.set(qn("w:val"), "808080")
        rpr.append(sz); rpr.append(col); r.append(rpr)
        if ftype:
            fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), ftype); r.append(fc)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve"); it.text = f" {content} "; r.append(it)
        para._p.append(r)


# ── Conversion PDF ────────────────────────────────────────────────────────────
def _to_pdf(docx_path: Path, pdf_path: Path):
    env = {**os.environ, "HOME": "/tmp", "DISPLAY": ""}
    cmd = [
        settings.LIBREOFFICE_PATH,
        "--headless", "--norestore", "--nofirststartwizard",
        "--convert-to", "pdf",
        "--outdir", str(pdf_path.parent),
        str(docx_path),
    ]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=180, env=env)
        logger.info(f"LibreOffice stdout: {r.stdout[:300]}")
        if r.stderr:
            logger.warning(f"LibreOffice stderr: {r.stderr[:300]}")
        generated = pdf_path.parent / (docx_path.stem + ".pdf")
        if generated.exists() and generated != pdf_path:
            generated.rename(pdf_path)
        if not pdf_path.exists():
            raise FileNotFoundError("PDF non genere")
        logger.info(f"PDF OK : {pdf_path.stat().st_size // 1024} Ko")
    except subprocess.TimeoutExpired:
        raise RuntimeError("LibreOffice timeout 180s")
    except FileNotFoundError:
        raise RuntimeError(f"LibreOffice introuvable : {settings.LIBREOFFICE_PATH}")


def _count_pages(doc) -> int:
    total = sum(max(1, len(p.text) // 80) for p in doc.paragraphs if p.text.strip())
    return max(1, total // 40)
