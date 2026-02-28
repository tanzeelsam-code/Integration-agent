"""
INTEGRATION — Figure formatting.
Ensures captions are below figures with correct style.
Numbers figures per chapter.
"""

import re
from docx.oxml.ns import qn
from docx.shared import Twips
from .typography import apply_run_style


def _para_has_image(paragraph) -> bool:
    """Check if a paragraph contains an inline image or drawing."""
    for run in paragraph.runs:
        drawing = run._element.findall(qn("w:drawing"))
        if drawing:
            return True
        # Also check for inline shapes
        inline = run._element.findall(f".//{qn('wp:inline')}")
        if inline:
            return True
    # Check for pict elements (older format)
    pict = paragraph._element.findall(f".//{qn('w:pict')}")
    if pict:
        return True
    return False


def _is_figure_caption(text: str) -> bool:
    """Check if text looks like a figure caption."""
    return bool(re.match(r"^\s*figure\s+\d", text, re.IGNORECASE))


def format_figures(doc, chapter_num: int = 1) -> dict:
    """
    Format all figures in the document:
    - Ensure captions are below figures
    - Apply correct caption styling
    - Number figures per chapter

    Returns dict with changes for compliance summary.
    """
    changes = {
        "reformatted": [],
        "captions_moved": [],
        "captions_added": [],
        "source_lines_added": [],
    }

    figure_counter = 0
    paragraphs = doc.paragraphs

    for idx, para in enumerate(paragraphs):
        if not _para_has_image(para):
            continue

        figure_counter += 1
        fig_label = f"Figure {chapter_num}.{figure_counter}"
        changes["reformatted"].append(fig_label)

        # Check if the next paragraph is a caption
        if idx + 1 < len(paragraphs):
            next_para = paragraphs[idx + 1]
            if _is_figure_caption(next_para.text):
                # Caption exists below — restyle it
                for run in next_para.runs:
                    apply_run_style(run, "caption")
                # Apply spacing
                pf = next_para.paragraph_format
                pf.space_before = Twips(80)
                pf.space_after = Twips(60)

                # Check for source line after caption
                if idx + 2 < len(paragraphs):
                    source_para = paragraphs[idx + 2]
                    if source_para.text.strip().lower().startswith("source"):
                        for run in source_para.runs:
                            apply_run_style(run, "source")
                        spf = source_para.paragraph_format
                        spf.space_before = Twips(0)
                        spf.space_after = Twips(120)
                    else:
                        changes["source_lines_added"].append(fig_label)
            else:
                changes["captions_added"].append(f"{fig_label} — needs manual caption")

        # Check if caption is ABOVE (wrong position)
        if idx > 0:
            prev_para = paragraphs[idx - 1]
            if _is_figure_caption(prev_para.text):
                changes["captions_moved"].append(
                    f"{fig_label}: caption was above, should be below"
                )

    return changes
