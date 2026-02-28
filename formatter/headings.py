"""
INTEGRATION — Heading detection, hierarchy, and auto-numbering.
Decimal multi-level numbering: 1. / 1.1. / 1.1.1. / 1.1.1.1.
"""

import re
from typing import Optional, List
from docx.shared import Pt, Twips
from docx.oxml.ns import qn
from . import constants as C
from .typography import apply_run_style


def classify_heading_level(paragraph) -> Optional[int]:
    """Return heading level (1-4) if paragraph is a heading, else None."""
    style_name = paragraph.style.name if paragraph.style else ""
    if style_name in C.HEADING_STYLES:
        return C.HEADING_STYLES[style_name]
    # Heuristic: detect by font size + bold
    for run in paragraph.runs:
        if run.font.bold:
            sz = run.font.size
            if sz:
                pt = sz.pt
                if pt >= 14:
                    return 1
                elif pt >= 12:
                    return 2
                elif pt >= 11:
                    return 3
    return None


def strip_existing_number(text: str) -> str:
    """Remove leading decimal numbering like '1.2.3.  Title' → 'Title'."""
    return re.sub(r"^[\d]+(?:\.[\d]+)*\.?\s+", "", text).strip()


def build_heading_number(counters: List[int]) -> str:
    """Build a decimal heading number string from counters, e.g. '1.2.3.'."""
    return ".".join(str(c) for c in counters) + "."


def renumber_headings(doc) -> list[dict]:
    """
    Walk all paragraphs, detect headings, renumber them with decimal
    multi-level numbering, apply correct typography and spacing.
    Returns a list of changes for the compliance summary.
    """
    counters = [0, 0, 0, 0]   # H1..H4
    changes = []

    for para in doc.paragraphs:
        level = classify_heading_level(para)
        if level is None:
            continue

        # Update counters
        counters[level - 1] += 1
        # Reset child counters
        for i in range(level, 4):
            counters[i] = 0

        active = counters[:level]
        new_number = build_heading_number(active)

        # Get the raw text (strip any existing numbering)
        raw_text = strip_existing_number(para.text)
        old_text = para.text
        new_text = f"{new_number}  {raw_text}"

        if old_text != new_text:
            changes.append({"old": old_text, "new": new_text, "level": level})

        # Rewrite paragraph: clear runs, add single run with correct text
        for run in para.runs:
            run._element.getparent().remove(run._element)

        new_run = para.add_run(new_text)

        # Determine element type key
        type_key = f"heading{level}"
        apply_run_style(new_run, type_key)

        # Set paragraph style
        style_name = C.LEVEL_TO_STYLE.get(level)
        if style_name:
            para.style = doc.styles[style_name] if style_name in [s.name for s in doc.styles] else para.style

        # Apply spacing
        spacing = C.HEADING_SPACING.get(level, {})
        pf = para.paragraph_format
        if spacing.get("before"):
            pf.space_before = Twips(spacing["before"])
        if spacing.get("after"):
            pf.space_after = Twips(spacing["after"])

    return changes
