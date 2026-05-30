# -*- coding: utf-8 -*-
"""
docx_template_builder.py — DOCX template-mode builder.

Mirrors the proven PPTX template-mode: open the SOURCE .docx and modify it
IN PLACE so every design element is preserved:
  - theme (theme1.xml: fonts + scheme colors)
  - cover page styling
  - existing tables and their colored fills
  - headers/footers, margins, default font

It then applies targeted, grounded edits:
  - spelling / consistency fixes (run-level, formatting preserved)
  - inserts/refreshes a RACI matrix table (activities x stakeholders)
  - inserts/refreshes a vertical process-flow table (stakeholder | step | ▼)

Brand colors are EXTRACTED from the source (dominant table-header fill +
theme accent), so the generated tables stay on-brand for ANY input document.
"""

import logging
from collections import Counter
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

logger = logging.getLogger(__name__)

# ── Sensible fallback palette (used only if extraction finds nothing) ─────────
_FALLBACK_BRAND   = "A91228"   # header / primary
_FALLBACK_ACCENT  = "1F3864"   # R/A accent (navy)
_GRAY_C           = "7B7B7B"   # Consulted
_GRAY_I           = "4A4A4A"   # Informed
_WHITE            = "FFFFFF"

# RACI code → fill color. R/A uses accent; C/I use grays; blank/'-' stays white.
def _raci_palette(brand: str, accent: str) -> dict:
    return {
        "R/A": accent, "A/R": accent,
        "R":   accent,
        "A":   brand,
        "C":   _GRAY_C,
        "I":   _GRAY_I,
        "-":   _WHITE, "":  _WHITE,
    }


# ── Public entry point ────────────────────────────────────────────────────────

def build_docx_from_template(
    source_path: str,
    output_path: str,
    *,
    spelling_fixes: dict | None = None,
    raci: dict | None = None,
    flow: dict | None = None,
) -> str:
    """
    Open the source .docx, modify in place, save to output_path.

    Args:
        source_path: path to the source .docx (its theme/styling is preserved)
        output_path: where to save the result
        spelling_fixes: {wrong: right} run-level text replacements
        raci: {
            "activities": ["1. Vendor Empanelment", ...],
            "stakeholders": ["SRM Team", "Procurement", ...],
            "grid": [["R/A","C",...], ...]   # one row per activity, one col per stakeholder
        }
        flow: {
            "title": "START – Vendor Empanelment Request",
            "steps": [{"stakeholder": "SRM Team", "description": "Initial Evaluation ..."}, ...]
        }

    Returns: absolute path to saved file.
    """
    doc = Document(source_path)

    brand, accent = _extract_brand_colors(doc)
    base_font = _extract_default_font(doc)
    logger.info(f"[DocxTemplate] brand={brand} accent={accent} font={base_font}")

    if spelling_fixes:
        n = _apply_spelling_fixes(doc, spelling_fixes)
        logger.info(f"[DocxTemplate] Applied {n} spelling/consistency fix(es)")

    if raci and raci.get("activities") and raci.get("stakeholders"):
        ok = _insert_raci_after_heading(doc, raci, brand, accent, base_font)
        logger.info(f"[DocxTemplate] RACI matrix inserted={ok}")

    if flow and flow.get("steps"):
        ok = _insert_flow_after_heading(doc, flow, brand, accent, base_font)
        logger.info(f"[DocxTemplate] Process flow inserted={ok}")

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    logger.info(f"[DocxTemplate] Saved → {out}")
    return str(out.resolve())


# ── Brand / font extraction ─────────────────────────────────────────────────

def _extract_brand_colors(doc) -> tuple[str, str]:
    """
    Determine the source's brand (header) color and an accent color.

    brand  = most common non-white/non-auto cell fill across existing tables
    accent = theme dk2 (or a navy fallback), distinct from brand
    """
    fills: Counter = Counter()
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                f = _get_cell_fill(c)
                if f and f.upper() not in ("FFFFFF", "AUTO", "FFFFF", "F2F2F2", "D9E1F2"):
                    fills[f.upper()] += 1

    brand = fills.most_common(1)[0][0] if fills else _FALLBACK_BRAND

    # Accent: prefer a dark navy from the theme color scheme; else fallback.
    accent = _theme_color(doc, "dk2") or _FALLBACK_ACCENT
    if accent.upper() == brand.upper():
        accent = _FALLBACK_ACCENT if brand.upper() != _FALLBACK_ACCENT else "0E2841"
    return brand, accent


def _theme_color(doc, key: str) -> str | None:
    """Read a theme scheme color (dk1/dk2/accent1...) from theme1.xml."""
    try:
        part = doc.part.package.part_related_by  # not used; read via zip below
    except Exception:
        pass
    try:
        import zipfile, re
        src = doc.part.package._partname  # unreliable; fall back to scanning parts
    except Exception:
        src = None
    # Most reliable: scan the theme part bytes already loaded in the package.
    try:
        for rel in doc.part.package.iter_parts():
            if rel.partname.endswith("theme1.xml") or "theme/theme" in str(rel.partname):
                xml = rel.blob.decode("utf-8", "ignore")
                import re
                m = re.search(rf"<a:{key}>\s*<a:srgbClr val=\"([0-9A-Fa-f]{{6}})\"", xml)
                if m:
                    return m.group(1).upper()
                m = re.search(rf"<a:{key}>\s*<a:sysClr[^>]*lastClr=\"([0-9A-Fa-f]{{6}})\"", xml)
                if m:
                    return m.group(1).upper()
    except Exception as e:
        logger.debug(f"[DocxTemplate] theme color read failed: {e}")
    return None


def _extract_default_font(doc) -> str:
    try:
        f = doc.styles["Normal"].font.name
        if f:
            return f
    except Exception:
        pass
    return "Arial"


# ── Spelling fixes (run-level, formatting preserved) ──────────────────────────

def _apply_spelling_fixes(doc, fixes: dict) -> int:
    count = 0

    def fix_paragraph(p):
        nonlocal count
        for run in p.runs:
            txt = run.text
            if not txt:
                continue
            new = txt
            for wrong, right in fixes.items():
                if wrong and wrong in new:
                    new = new.replace(wrong, right)
            if new != txt:
                run.text = new
                count += 1

    for p in doc.paragraphs:
        fix_paragraph(p)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    fix_paragraph(p)
    return count


# ── RACI matrix insertion ─────────────────────────────────────────────────────

def _insert_raci_after_heading(doc, raci: dict, brand: str, accent: str, base_font: str) -> bool:
    activities   = raci["activities"]
    stakeholders = raci["stakeholders"]
    grid         = raci.get("grid") or []

    heading = _find_heading(doc, ["raci matrix", "raci"])
    if heading is None:
        heading = _append_heading(doc, "RACI Matrix")

    _remove_following_table(heading, matches=_looks_like_raci)

    n_rows = len(activities) + 1
    n_cols = len(stakeholders) + 1
    tbl = doc.add_table(rows=n_rows, cols=n_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    palette = _raci_palette(brand, accent)

    # Header row
    hdr = tbl.rows[0].cells
    _style_cell(hdr[0], "Activity", fill=brand, color=_WHITE, bold=True, size=9, font=base_font)
    for ci, sh in enumerate(stakeholders, start=1):
        _style_cell(hdr[ci], sh, fill=brand, color=_WHITE, bold=True, size=9, font=base_font,
                    align="center")

    # Body
    for ri, activity in enumerate(activities, start=1):
        cells = tbl.rows[ri].cells
        _style_cell(cells[0], activity, fill=_WHITE, color="000000", bold=False, size=9, font=base_font)
        row_codes = grid[ri - 1] if ri - 1 < len(grid) else []
        for ci in range(1, n_cols):
            code = (row_codes[ci - 1] if ci - 1 < len(row_codes) else "-") or "-"
            code = str(code).strip().upper().replace(" ", "")
            code = {"RA": "R/A", "AR": "A/R"}.get(code, code)
            fill = palette.get(code, _WHITE)
            is_filled = fill != _WHITE
            _style_cell(
                cells[ci], code if code not in ("", "-") else "-",
                fill=fill, color=_WHITE if is_filled else "000000",
                bold=is_filled, size=9, font=base_font, align="center",
            )

    _set_table_borders(tbl, "BFBFBF")
    _set_raci_col_widths(tbl, n_cols)
    _move_table_after(tbl, heading)
    _add_raci_legend(tbl, heading, palette, base_font)
    return True


def _set_raci_col_widths(tbl, n_cols):
    """First column wide (~1.75"), code columns narrow (~0.53")."""
    first = Emu(1600200)
    rest = Emu(482600)
    for row in tbl.rows:
        row.cells[0].width = first
        for ci in range(1, n_cols):
            row.cells[ci].width = rest


def _add_raci_legend(tbl, heading, palette, base_font):
    """Add a one-line legend paragraph right after the RACI table."""
    legend = OxmlElement("w:p")
    tbl._tbl.addnext(legend)
    from docx.text.paragraph import Paragraph
    p = Paragraph(legend, tbl._parent)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    parts = [
        ("R = Responsible   ", "000000"),
        ("A = Accountable   ", "000000"),
        ("C = Consulted   ", "000000"),
        ("I = Informed   ", "000000"),
        ("R/A = Responsible & Accountable", "000000"),
    ]
    for text, col in parts:
        run = p.add_run(text)
        run.font.size = Pt(8)
        run.font.italic = True
        if base_font:
            run.font.name = base_font
        run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)


# ── Process-flow insertion ─────────────────────────────────────────────────────

def _insert_flow_after_heading(doc, flow: dict, brand: str, accent: str, base_font: str) -> bool:
    steps = flow["steps"]
    title = flow.get("title", "START")

    heading = _find_heading(
        doc, ["end-to-end process flow", "end-to-end flow", "process flow", "process workflow"]
    )
    if heading is None:
        heading = _append_heading(doc, "Vendor Empanelment End-to-End Process Flow")

    _remove_following_table(heading, matches=_looks_like_flow)

    # Layout: header row + per step (content row + arrow row), minus trailing arrow
    n_rows = 1 + len(steps) * 2 - 1
    tbl = doc.add_table(rows=n_rows, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header
    h = tbl.rows[0].cells
    _style_cell(h[0], "Stakeholder", fill=brand, color=_WHITE, bold=True, size=8, font=base_font, align="center")
    _style_cell(h[1], title, fill=accent, color=_WHITE, bold=True, size=9, font=base_font, align="center")
    _style_cell(h[2], "", fill=_WHITE, color="000000", bold=False, size=8, font=base_font)

    r = 1
    for i, step in enumerate(steps):
        stk = str(step.get("stakeholder", "")).strip()
        desc = str(step.get("description", "")).strip()
        cells = tbl.rows[r].cells
        _style_cell(cells[0], stk, fill=brand, color=_WHITE, bold=True, size=8, font=base_font, align="center")
        _style_cell(cells[1], desc, fill=accent if i == len(steps) - 1 else brand,
                    color=_WHITE, bold=True, size=9, font=base_font)
        _style_cell(cells[2], "", fill=_WHITE, color="000000", bold=False, size=8, font=base_font)
        r += 1
        # arrow row (skip after last step)
        if i < len(steps) - 1:
            acells = tbl.rows[r].cells
            _style_cell(acells[0], "", fill=_WHITE, color="000000", bold=False, size=8, font=base_font)
            _style_cell(acells[1], "▼", fill=_WHITE, color=brand, bold=True, size=12, font=base_font, align="center")
            _style_cell(acells[2], "", fill=_WHITE, color="000000", bold=False, size=8, font=base_font)
            r += 1

    _set_table_borders(tbl, "FFFFFF")  # flow uses fills, light/no visible borders
    _set_flow_col_widths(tbl)
    _move_table_after(tbl, heading)
    return True


def _set_flow_col_widths(tbl):
    w0, w1, w2 = Emu(914400), Emu(4114800), Emu(914400)
    for row in tbl.rows:
        row.cells[0].width = w0
        row.cells[1].width = w1
        row.cells[2].width = w2


# ── Low-level docx helpers ─────────────────────────────────────────────────────

def _style_cell(cell, text, *, fill, color, bold, size, font, align="left"):
    """Set a cell's text + fill + font in one shot, clearing prior content.

    Produces exactly ONE formatted run (no stray empty runs)."""
    _set_cell_fill(cell, fill)
    # remove extra paragraphs, keep the first
    while len(cell.paragraphs) > 1:
        pe = cell.paragraphs[-1]._p
        pe.getparent().remove(pe)
    p = cell.paragraphs[0]
    # remove all existing runs from the first paragraph
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    run = p.add_run(str(text))
    run.font.size = Pt(size)
    run.font.bold = bold
    if font:
        run.font.name = font
    rgb = _hex_rgb(color)
    if rgb:
        run.font.color.rgb = RGBColor(*rgb)
    # tighten cell margins for compact tables
    _set_cell_margins(cell, top=40, bottom=40, left=60, right=60)


def _set_cell_fill(cell, hex_color: str):
    if not hex_color:
        return
    tcPr = cell._tc.get_or_add_tcPr()
    for shd in tcPr.findall(qn("w:shd")):
        tcPr.remove(shd)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.lstrip("#").upper())
    tcPr.append(shd)


def _set_cell_margins(cell, *, top, bottom, left, right):
    tcPr = cell._tc.get_or_add_tcPr()
    existing = tcPr.find(qn("w:tcMar"))
    if existing is not None:
        tcPr.remove(existing)
    tcMar = OxmlElement("w:tcMar")
    for tag, val in (("top", top), ("bottom", bottom), ("start", left), ("end", right),
                     ("left", left), ("right", right)):
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)


def _set_table_borders(tbl, hex_color: str):
    tblPr = tbl._tbl.tblPr
    existing = tblPr.find(qn("w:tblBorders"))
    if existing is not None:
        tblPr.remove(existing)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), hex_color.lstrip("#").upper())
        borders.append(el)
    tblPr.append(borders)


def _get_cell_fill(cell) -> str | None:
    tcPr = cell._tc.tcPr
    if tcPr is not None:
        shd = tcPr.find(qn("w:shd"))
        if shd is not None:
            f = shd.get(qn("w:fill"))
            if f and f != "auto":
                return f
    return None


def _find_heading(doc, keywords: list[str]):
    """
    Find the heading paragraph for a real body section matching any keyword.

    Documents often contain a manually-typed Table of Contents whose entries
    are ALSO styled as headings and match the same keywords. The TOC sits near
    the top, the real section near the bottom — so we return the LAST matching
    heading-styled paragraph (falling back to the last plain match).
    """
    kws = [k.lower() for k in keywords]
    last_heading = None
    last_any = None
    for p in doc.paragraphs:
        txt = p.text.strip().lower()
        if not txt or len(txt) >= 80:
            continue
        style = (p.style.name or "").lower() if p.style else ""
        is_heading = "heading" in style or "title" in style
        if any(k in txt for k in kws):
            last_any = p
            if is_heading:
                last_heading = p
    return last_heading or last_any


def _append_heading(doc, text: str):
    """Append a Heading-4-styled paragraph at the end (page break before)."""
    p = doc.add_paragraph()
    try:
        p.style = doc.styles["Heading 4"]
    except Exception:
        try:
            p.style = doc.styles["Heading 1"]
        except Exception:
            pass
    run = p.add_run(text)
    run.bold = True
    return p


def _is_heading_p(p_el) -> bool:
    """True if a <w:p> element is styled as a Heading/Title paragraph."""
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return False
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        return False
    val = (pStyle.get(qn("w:val")) or "").lower()
    return val.startswith("heading") or val.startswith("title")


def _table_text(tbl_el) -> str:
    """Concatenate all text in a <w:tbl> element (lowercased)."""
    texts = [t.text or "" for t in tbl_el.iter(qn("w:t"))]
    return " ".join(texts).lower()


def _looks_like_raci(tbl_el) -> bool:
    """A RACI grid: first header cell 'activity' OR many standalone R/A/C/I cells."""
    cells = [(t.text or "").strip() for t in tbl_el.iter(qn("w:t"))]
    if cells and cells[0].lower() in ("activity", "activities"):
        return True
    codes = sum(1 for c in cells if c.upper() in ("R", "A", "C", "I", "R/A", "A/R"))
    return codes >= 6


def _looks_like_flow(tbl_el) -> bool:
    """A process-flow table: contains ▼ arrows OR a 'stakeholder' header column."""
    txt = _table_text(tbl_el)
    if "▼" in txt:                      # ▼
        return True
    cells = [(t.text or "").strip().lower() for t in tbl_el.iter(qn("w:t"))]
    return bool(cells) and cells[0] == "stakeholder"


def _remove_following_table(heading_para, matches=None):
    """
    Remove the table(s) belonging to this heading's section.

    Scans forward to the next heading, but removes a <w:tbl> ONLY when the
    `matches(tbl_el)` predicate returns True. This guarantees we never delete an
    unrelated table (e.g. a definitions table) that happens to share a section.
    If `matches` is None, removes the first table found (legacy behavior).
    """
    nxt = heading_para._p.getnext()
    removed = False
    while nxt is not None:
        tag = nxt.tag.split("}")[-1]
        if tag == "tbl":
            following = nxt.getnext()
            if matches is None or matches(nxt):
                nxt.getparent().remove(nxt)
                removed = True
            nxt = following
            continue
        if tag == "p":
            if _is_heading_p(nxt):
                break          # reached the next section
            nxt = nxt.getnext()
            continue
        break                   # sectPr or other structural element
    return removed


def _move_table_after(tbl, heading_para):
    """Move a table element to sit right after the heading paragraph."""
    heading_para._p.addnext(tbl._tbl)


def _hex_rgb(hex_color: str):
    if not hex_color:
        return None
    h = str(hex_color).lstrip("#").strip()
    if len(h) != 6:
        return None
    try:
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except ValueError:
        return None
