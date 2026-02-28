"""
INTEGRATION Energy Plus — Formatting Constants
All colour hex values, typography specs, page layout dimensions,
bullet list specs, and DXA spacing values.
"""

# ─── COLOUR PALETTE (exact hex) ───────────────────────────────────────
GREEN_HAZE   = "009959"
WHITE        = "FFFFFF"
ALT_ROW      = "E2EFD9"
BODY_TEXT     = "000000"
SOURCE_TEXT   = "404040"
HEADER_GREY  = "606060"
INNER_BORDER = "80CCB0"
HYPERLINK    = "0563C1"

# ─── FONT FAMILY ──────────────────────────────────────────────────────
FONT_FAMILY = "Arial"
DOCUMENT_LANG = "en-PK"

# ─── TYPOGRAPHY SPECS ─────────────────────────────────────────────────
# (size_pt, bold, italic, colour_hex)
TYPO = {
    "cover_title":     (18, True,  False, GREEN_HAZE),
    "cover_subtitle":  (13, True,  False, BODY_TEXT),
    "heading1":        (14, True,  False, GREEN_HAZE),
    "heading2":        (12, True,  False, GREEN_HAZE),
    "heading3":        (11, True,  False, GREEN_HAZE),
    "heading4":        (11, True,  True,  GREEN_HAZE),
    "body":            (11, False, False, BODY_TEXT),
    "bullet":          (12, False, False, BODY_TEXT),
    "caption":         (10, True,  False, BODY_TEXT),
    "source":          (9,  False, True,  SOURCE_TEXT),
    "footnote":        (10, False, False, BODY_TEXT),
    "header":          (9,  False, False, HEADER_GREY),
    "footer_label":    (9,  False, False, HEADER_GREY),
    "footer_page":     (9,  True,  False, GREEN_HAZE),
}

# ─── PAGE LAYOUT ──────────────────────────────────────────────────────
PAGE_WIDTH_MM     = 210
PAGE_HEIGHT_MM    = 297
PAGE_WIDTH_DXA    = 11906   # 210mm in DXA (twips)
PAGE_HEIGHT_DXA   = 16838   # 297mm in DXA
MARGIN_DXA        = 1134    # 2 cm in DXA
HEADER_DIST_DXA   = 708
FOOTER_DIST_DXA   = 708

# ─── PARAGRAPH SPACING ───────────────────────────────────────────────
BODY_SPACE_BEFORE  = 120   # 6 pt ≈ 120 DXA
BODY_SPACE_AFTER   = 120
LINE_SPACING       = 1.08  # single (1.08)

# ─── HEADING SPACING (DXA) ───────────────────────────────────────────
HEADING_SPACING = {
    1: {"before": 320, "after": 160},
    2: {"before": 240, "after": 120},
    3: {"before": 200, "after": 80},
    4: {"before": 160, "after": 60},
}

# ─── TABLE CONSTANTS ─────────────────────────────────────────────────
TABLE_WIDTH_DXA         = 9000
TABLE_CELL_PAD_TOP      = 80
TABLE_CELL_PAD_BOTTOM   = 80
TABLE_CELL_PAD_LEFT     = 120
TABLE_CELL_PAD_RIGHT    = 120
TABLE_HEADER_BORDER_SZ  = 8    # eighth-points (8 = 1pt visual)
TABLE_OUTER_BORDER_SZ   = 6
TABLE_INNER_BORDER_SZ   = 4

# Caption spacing
TABLE_CAPTION_BEFORE    = 120
TABLE_CAPTION_AFTER     = 60
TABLE_SOURCE_BEFORE     = 40
TABLE_SOURCE_AFTER      = 120

# ─── FIGURE CONSTANTS ────────────────────────────────────────────────
FIG_CAPTION_BEFORE      = 80
FIG_CAPTION_AFTER       = 60
FIG_SOURCE_BEFORE       = 0
FIG_SOURCE_AFTER        = 120

# ─── HEADER / FOOTER ─────────────────────────────────────────────────
HEADER_FOOTER_BORDER_SZ = 8
FOOTER_TAB_STOP_DXA     = 9000

# Header content template
HEADER_TEMPLATE = "Energy Plus: Energy Master Plan for Gilgit-Baltistan    |    {report_name}"
FOOTER_LEFT_TEMPLATE = "© {year} INTEGRATION  |  {project_number}"

# ─── BULLET LIST ──────────────────────────────────────────────────────
BULLET_LEVELS = [
    {"char": "\u25CF", "indent": 720,  "hanging": 360},  # Level 0: ●
    {"char": "\u25CB", "indent": 1440, "hanging": 360},  # Level 1: ○
    {"char": "\u25A0", "indent": 2160, "hanging": 360},  # Level 2: ■
]

# ─── HEADING STYLE NAMES (python-docx built-in style names) ──────────
HEADING_STYLES = {
    "Heading 1": 1,
    "Heading 2": 2,
    "Heading 3": 3,
    "Heading 4": 4,
}

# Reverse: level -> style name
LEVEL_TO_STYLE = {v: k for k, v in HEADING_STYLES.items()}
