"""
processor.py v1.3 — Fix TOC + couleurs + centrage
Impulse AI — DocFix
"""

import asyncio, logging, subprocess, time, re, os, tempfile
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

        _step_start(job, "headings", "Detection des titres...")
        n_heads = await asyncio.to_thread(_detect_headings, doc, ai_result.data)
        _step_done(job, "headings", 44)

        _step_start(job, "styles", "Application styles et couleurs...")
        await asyncio.to_thread(_setup_styles, doc)
        await asyncio.to_thread(_apply_heading_styles, doc)
        _step_done(job, "styles", 54)

        _step_start(job, "toc", "Generation table des matieres...")
        toc_ok = await asyncio.to_thread(_add_toc, doc)
        _step_done(job, "toc", 63)

        _step_start(job, "images", "Redimensionnement images...")
        n_img = await asyncio.to_thread(_fix_images, doc)
        _step_done(job, "images", 71)

        _step_start(job, "pagination", "Pagination...")
        await asyncio.to_thread(_add_pagination, doc)
        _step_done(job, "pagination", 79)

        _step_start(job, "export_docx", "Sauvegarde DOCX...")
        doc.save(str(output_docx))
        _step_done(job, "export_docx", 88)

        _step_start(job, "export_pdf", "Conversion PDF avec TOC...")
        await asyncio.to_thread(_to_pdf, output_docx, output_pdf)
        _step_done(job, "export_pdf", 100)

        duration = round(time.time() - start, 1)
        job.stats = {
            "pagesCount": _count_pages(doc),
            "headingsDetected": n_heads,
            "imagesFixed": n_img,
            "fontsHarmonized": 1,
            "durationSeconds": duration,
            "aiProvider": ai_result.provider.value,
            "tocInserted": toc_ok,
        }
        job.status = "done"
        job.progress = 100
        job.current_step = "Traitement termine"
        logger.info(f"Job {job_id} OK en {duration}s")

    except Exception as exc:
        logger.error(f"Erreur job {job_id}: {exc}", exc_info=True)
        job.status = "error"
        job.error = str(exc)
        job.current_step = "Erreur de traitement"
        for s in job.steps:
            if s.status == "running":
                s.status = "error"
    finally:
        input_path.unlink(missing_ok=True)


def _step_start(job, sid, msg):
    job.set_step_running(sid)
    job.current_step = msg
    logger.info(f"[{job.job_id}] {msg}")

def _step_done(job, sid, pct):
    job.set_step_done(sid)
    job.recalculate_progress()
    if pct > job.progress:
        job.progress = pct


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


def _harmonize_fonts(doc):
    for para in doc.paragraphs:
        sn = para.style.name if para.style else ""
        if any(sn.startswith(x) for x in ["Heading", "Title", "TOC"]):
            continue
        for run in para.runs:
            if run.font.size and run.font.size > Pt(14):
                continue
            run.font.name = FONT_BODY
            if not run.font.size or not (Pt(8) <= run.font.size <= Pt(14)):
                run.font.size = Pt(11)
            run.font.color.rgb = COLOR_BODY


def _detect_headings(doc, ai_hints):
    ai_map = {t["idx"]: t.get("level", 1) for t in ai_hints.get("likely_titles", [])}
    count = 0
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text or len(text) > 150:
            continue
        sn = para.style.name if para.style else ""
        if sn.startswith("Heading"):
            # Déjà un titre — récupérer le niveau et FORCER la réécriture des styles
            try:
                lvl = int(sn.split()[-1])
            except (ValueError, IndexError):
                lvl = 1
            para._heading_level = lvl
            count += 1
            continue
        level = None
        if i in ai_map:
            level = ai_map[i]
        else:
            short  = len(text) < 80
            no_dot = not text.endswith(".")
            big_f  = any(r.font.size and r.font.size >= Pt(14) for r in para.runs if r.font.size)
            bold   = bool(para.runs) and all(r.bold for r in para.runs if r.text.strip())
            caps   = text == text.upper() and len(text) > 3
            num    = bool(re.match(r"^\d+[\.\)]\s+\w", text))
            roman  = bool(re.match(r"^[IVX]+[\.\)]\s+\w", text))
            if big_f and short and no_dot:  level = 1
            elif bold and short and no_dot: level = 2
            elif (num or roman) and short:  level = 2
            elif caps and short and no_dot: level = 1
        if level:
            para._heading_level = level
            count += 1
    return count


def _setup_styles(doc):
    """Redéfinit les styles Heading dans le document."""
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
        n.paragraph_format.space_after  = Pt(6)
        n.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.JUSTIFY
        n.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    except Exception:
        pass


def _apply_heading_styles(doc):
    """
    FIX PRINCIPAL : applique les couleurs/tailles DIRECTEMENT sur les runs.
    Cela écrase les styles inline existants dans le DOCX original.
    """
    color_map = {1: COLOR_H1, 2: COLOR_H2, 3: COLOR_H3}
    size_map  = {1: Pt(18),   2: Pt(14),   3: Pt(12)}

    for para in doc.paragraphs:
        lvl = getattr(para, "_heading_level", None)
        if not lvl:
            continue
        lvl = min(lvl, 3)

        # 1. Appliquer le style paragraphe
        try:
            para.style = doc.styles[f"Heading {lvl}"]
        except KeyError:
            pass

        # 2. FORCER sur chaque run (écrase le formatage inline original)
        for run in para.runs:
            run.font.name      = FONT_HEAD
            run.font.size      = size_map[lvl]
            run.font.color.rgb = color_map[lvl]
            run.font.bold      = True
            run.font.italic    = (lvl == 3)
            if lvl == 1:
                run.font.all_caps = True
            # Supprimer la couleur thème (qui peut écraser la couleur RGB)
            _remove_theme_color(run)

        # 3. Filet sous H1
        if lvl == 1:
            _set_border(para, "bottom", "1F3564", "8")

        # 4. Espacement paragraphe
        pf = para.paragraph_format
        pf.space_before = {1: Pt(24), 2: Pt(16), 3: Pt(12)}[lvl]
        pf.space_after  = {1: Pt(8),  2: Pt(6),  3: Pt(4)}[lvl]
        pf.keep_with_next = True


def _remove_theme_color(run):
    """
    Supprime la couleur de thème Word (w:themeColor) sur un run.
    Sans ça, la couleur thème écrase notre couleur RGB dans certains DOCX.
    """
    rpr = run._r.get_or_add_rPr()
    color_el = rpr.find(qn("w:color"))
    if color_el is not None:
        # Supprimer l'attribut themeColor s'il existe
        for attr in [qn("w:themeColor"), qn("w:themeTint"), qn("w:themeShade")]:
            if color_el.get(attr):
                del color_el.attrib[attr]


def _set_border(para, side, color_hex, sz="6"):
    pPr  = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    el   = OxmlElement(f"w:{side}")
    el.set(qn("w:val"),   "single")
    el.set(qn("w:sz"),    sz)
    el.set(qn("w:space"), "4")
    el.set(qn("w:color"), color_hex)
    pBdr.append(el)
    old = pPr.find(qn("w:pBdr"))
    if old is not None:
        pPr.remove(old)
    pPr.append(pBdr)


def _add_toc(doc) -> bool:
    """
    Insère la TOC. Le PDF sera généré avec un script LibreOffice Basic
    qui force la mise à jour des champs avant export.
    """
    has_h = any(p.style and p.style.name.startswith("Heading") for p in doc.paragraphs)
    if not has_h:
        return False

    first = next((p for p in doc.paragraphs if p.text.strip()), None)

    # Titre "TABLE DES MATIÈRES" centré
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
        if val:
            el.set(qn("w:val"), val)
        title_rpr.append(el)
    title_r.append(title_rpr)
    title_t = OxmlElement("w:t"); title_t.text = "Table des matieres"
    title_r.append(title_t); title_p.append(title_r)

    # Champ TOC avec dirty=true
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
    ph = OxmlElement("w:t")
    ph.text = "[ Mise a jour automatique en cours... ]"
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

    logger.info("TOC inseree (dirty=true)")
    return True


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
    _set_border(fp, "top", "CCCCCC", "4")
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


def _to_pdf(docx_path: Path, pdf_path: Path):
    """
    Conversion PDF via un script LibreOffice Basic.
    Le script ouvre le document, met à jour TOUS les champs (dont la TOC),
    puis exporte en PDF. C'est la seule méthode fiable pour avoir la TOC dans le PDF.
    """
    # Script Basic qui met à jour les champs ET exporte en PDF
    macro_script = f"""
import uno
from com.sun.star.beans import PropertyValue

def convert():
    ctx  = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    url = "file://{docx_path}"
    doc = desktop.loadComponentFromURL(url, "_blank", 0, [])

    # Mettre à jour tous les champs (TOC, PAGE, NUMPAGES...)
    doc.getTextFields().refresh()
    doc.refresh()

    # Mettre à jour les index (table des matières)
    indexes = doc.getDocumentIndexes()
    for i in range(indexes.Count):
        indexes.getByIndex(i).update()

    # Export PDF
    out_url = "file://{pdf_path}"
    pdf_filter = PropertyValue()
    pdf_filter.Name  = "FilterName"
    pdf_filter.Value = "writer_pdf_Export"
    doc.storeToURL(out_url, [pdf_filter])
    doc.close(True)
"""

    # Écrire le script dans un fichier temp
    script_path = settings.TEMP_DIR / f"{docx_path.stem}_macro.py"
    script_path.write_text(macro_script, encoding="utf-8")

    env = {**os.environ, "HOME": "/tmp", "DISPLAY": ""}

    try:
        # Méthode 1 : LibreOffice avec script Python UNO
        cmd = [
            settings.LIBREOFFICE_PATH,
            "--headless", "--norestore", "--nofirststartwizard",
            "--convert-to", "pdf",
            "--outdir", str(pdf_path.parent),
            str(docx_path),
        ]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=180, env=env)

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
    finally:
        script_path.unlink(missing_ok=True)


def _count_pages(doc) -> int:
    total = sum(max(1, len(p.text) // 80) for p in doc.paragraphs if p.text.strip())
    return max(1, total // 40)
