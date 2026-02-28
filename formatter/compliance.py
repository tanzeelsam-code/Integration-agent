"""
INTEGRATION — Compliance summary generator.
Produces the structured compliance report of all changes made.
"""


def generate_summary(
    headings_changed: list[dict],
    tables_reformatted: list[str],
    figures_info: dict,
    caption_corrections: list[str],
    fonts_corrected: list[str],
    colours_corrected: list[str],
    spacing_corrected: int,
    header_applied: bool,
    footer_applied: bool,
    bullets_fixed: int,
    hyperlinks_fixed: int,
) -> str:
    """Generate the exact-format compliance summary string."""

    heading_lines = []
    for h in headings_changed:
        heading_lines.append(f"  H{h['level']}: \"{h['old']}\" → \"{h['new']}\"")
    headings_str = "\n".join(heading_lines) if heading_lines else "  None"

    tables_str = ", ".join(tables_reformatted) if tables_reformatted else "None"

    figs_reformatted = figures_info.get("reformatted", [])
    figs_str = ", ".join(figs_reformatted) if figs_reformatted else "None"

    captions_moved = figures_info.get("captions_moved", [])
    captions_added = figures_info.get("captions_added", [])
    all_caption_fixes = caption_corrections + captions_moved + captions_added
    captions_str = "; ".join(all_caption_fixes) if all_caption_fixes else "None"

    source_added = figures_info.get("source_lines_added", [])
    source_str = ", ".join(source_added) if source_added else "None"

    unique_fonts = list(set(fonts_corrected))
    fonts_str = ", ".join(unique_fonts) if unique_fonts else "None"

    colours_str = f"{len(colours_corrected)} non-compliant colours replaced" if colours_corrected else "None"

    spacing_str = f"{spacing_corrected} elements corrected" if spacing_corrected else "None"

    hf_issues = []
    if not header_applied:
        hf_issues.append("Header could not be applied")
    if not footer_applied:
        hf_issues.append("Footer could not be applied")
    hf_str = "Yes" if (header_applied and footer_applied) else "; ".join(hf_issues)

    missing = []
    for item in all_caption_fixes:
        if "missing" in item.lower() or "needs manual" in item.lower():
            missing.append(item)
    if source_added:
        missing.extend([f"Source line needed for {s}" for s in source_added])
    missing_str = "; ".join(missing) if missing else "None"

    summary = f"""COMPLIANCE SUMMARY
==================
Headings renumbered:      {len(headings_changed)} changed
{headings_str}
Tables reformatted:       {len(tables_reformatted)} — {tables_str}
Figures reformatted:      {len(figs_reformatted)} — {figs_str}
Captions corrected:       {captions_str}
Source lines added:        {source_str}
Fonts corrected:          {fonts_str}
Colours corrected:        {colours_str}
Spacing corrected:        {spacing_str}
Bullets reformatted:      {bullets_fixed}
Hyperlinks styled:        {hyperlinks_fixed}
Header/footer applied:    {hf_str}
Missing elements added:   {missing_str}"""

    return summary
