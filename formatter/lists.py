"""
INTEGRATION — Bullet list formatting.
Replaces non-standard bullets with correct Unicode characters at correct indents.
"""

from typing import Optional
from docx.shared import Pt, Twips
from docx.oxml.ns import qn
from . import constants as C
from .typography import apply_run_style


def _get_list_level(paragraph) -> Optional[int]:
    """Detect the list indent level of a paragraph (0-based), or None if not a list."""
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        return None

    numPr = pPr.find(qn("w:numPr"))
    if numPr is not None:
        ilvl = numPr.find(qn("w:ilvl"))
        if ilvl is not None:
            return int(ilvl.get(qn("w:val"), "0"))
        return 0

    # Heuristic: check for bullet-like characters at the start
    text = paragraph.text.strip()
    bullet_chars = ["\u2022", "\u25CF", "\u25CB", "\u25A0", "\u25AA", "-", "•", "●", "○", "■"]
    for bc in bullet_chars:
        if text.startswith(bc):
            # Determine level by indentation
            ind = pPr.find(qn("w:ind"))
            if ind is not None:
                left = int(ind.get(qn("w:left"), "0"))
                if left >= 2160:
                    return 2
                elif left >= 1440:
                    return 1
            return 0

    return None


def _is_bullet_paragraph(paragraph) -> bool:
    """Check if paragraph is a bullet list item."""
    return _get_list_level(paragraph) is not None


def format_bullet_lists(doc) -> int:
    """
    Format all bullet list items:
    - Replace bullets with correct Unicode characters
    - Set correct indentation per level
    - Apply Arial 12pt styling

    Returns count of bullets reformatted.
    """
    count = 0

    for para in doc.paragraphs:
        level = _get_list_level(para)
        if level is None:
            continue

        # Clamp level to 0-2
        level = min(level, 2)
        spec = C.BULLET_LEVELS[level]
        count += 1

        # Strip existing bullet character from text if present
        text = para.text.strip()
        bullet_chars = ["\u2022", "\u25CF", "\u25CB", "\u25A0", "\u25AA", "-", "•", "●", "○", "■"]
        for bc in bullet_chars:
            if text.startswith(bc):
                text = text[len(bc):].strip()
                break

        # Set indentation
        pf = para.paragraph_format
        pf.left_indent = Twips(spec["indent"])
        pf.first_line_indent = Twips(-spec["hanging"])

        # Apply typography to all runs
        for run in para.runs:
            apply_run_style(run, "bullet")

    return count
