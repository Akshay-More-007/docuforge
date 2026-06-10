"""Fingerprint the visual style of .docx files for comparison.

For each file: producer metadata, fonts used (run-level histogram), heading
styles, table count + header fills, RACI/flow signatures, image count.
Usage: python scripts/style_fingerprint.py FILE [FILE...]
"""

import sys
from collections import Counter
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def iter_block_items(doc):
    from docx.oxml.ns import qn

    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def cell_fill(cell):
    shd = cell._tc.find(".//w:shd", NS)
    if shd is not None:
        return shd.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill")
    return None


def fingerprint(path: str):
    doc = Document(path)
    print(f"\n{'=' * 70}\n{Path(path).name}")
    cp = doc.core_properties
    print(f"  modified={cp.modified}  last_modified_by={cp.last_modified_by!r}")

    fonts = Counter()
    font_colors = Counter()
    headings = Counter()
    tables = 0
    raci = flow = False
    header_fills = Counter()

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            style = block.style.name if block.style else "?"
            if style.lower().startswith("heading") or style.lower().startswith("title"):
                headings[style] += 1
            for run in block.runs:
                if not run.text.strip():
                    continue
                fonts[run.font.name or f"(style:{style})"] += len(run.text)
                if run.font.color and run.font.color.rgb:
                    font_colors[str(run.font.color.rgb)] += len(run.text)
        else:
            tables += 1
            rows = block.rows
            if not rows:
                continue
            hdr = [c.text.strip().lower() for c in rows[0].cells]
            joined = " ".join(hdr)
            if "activity" in joined and len(hdr) >= 4:
                raci = True
            body_text = " ".join(c.text for r in rows for c in r.cells)
            if "▼" in body_text or ("stakeholder" in joined and "step" in joined):
                flow = True
            for c in rows[0].cells:
                f = cell_fill(c)
                if f and f not in ("auto",):
                    header_fills[f] += 1

    imgs = len(doc.inline_shapes)
    print(f"  tables={tables}  images={imgs}  RACI={raci}  flow={flow}")
    print(f"  headings: {dict(headings)}")
    print(f"  fonts(top5 by chars): {fonts.most_common(5)}")
    print(f"  font colors(top5): {font_colors.most_common(5)}")
    print(f"  table header fills(top6): {header_fills.most_common(6)}")


if __name__ == "__main__":
    for p in sys.argv[1:]:
        try:
            fingerprint(p)
        except Exception as e:
            print(f"\n{p}: FAILED — {e}")
