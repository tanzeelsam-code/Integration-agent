# INTEGRATION Document Reformatting Agent

**Energy Plus: Energy Master Plan for Gilgit-Baltistan, Pakistan**

A Python-based tool that takes any `.docx` document and reformats it to be **fully compliant** with the INTEGRATION Formatting Style Guide.

---

## ✨ Features

- **Full Style Compliance** — Applies exact colour palette, typography, heading hierarchy, table/figure formatting, header/footer, and bullet list rules
- **Automatic Heading Numbering** — Detects and renumbers all headings with decimal multi-level numbering
- **Table Reformatting** — Green header rows, alternating fills, correct border colours & sizes
- **Figure Handling** — Validates caption placement and source lines
- **Header & Footer** — Automated INTEGRATION branding with page numbers
- **Compliance Summary** — Detailed change report after every reformatting run
- **Two Interfaces** — CLI for batch/scripting, Web UI for drag-and-drop ease

---

## 🚀 Quick Start

### Install Dependencies

```bash
cd integration-agent
pip install -r requirements.txt
```

### CLI Usage

```bash
python main.py input.docx -o output.docx

# With custom header/footer values:
python main.py input.docx -o output.docx \
    --report "Energy Master Plan" \
    --year 2026 \
    --project "EMP-GB-001"
```

### Web UI

```bash
python app.py
# Open http://localhost:5000 in your browser
```

---

## 📁 Project Structure

```
integration-agent/
├── main.py              # CLI entry point
├── app.py               # Flask web UI
├── formatter/
│   ├── constants.py     # Colour palette, typography, layout constants
│   ├── engine.py        # Core orchestrator (12-step pipeline)
│   ├── headings.py      # Heading detection & auto-numbering
│   ├── tables.py        # Table formatting (fills, borders, captions)
│   ├── figures.py       # Figure formatting (captions, source lines)
│   ├── typography.py    # Font enforcement (Arial everywhere)
│   ├── layout.py        # Page layout, header & footer
│   ├── lists.py         # Bullet list formatting
│   └── compliance.py    # Compliance summary generator
├── templates/
│   └── index.html       # Web UI template
├── requirements.txt
└── README.md
```

---

## 🎨 INTEGRATION Style Guide Summary

| Element | Specification |
|---|---|
| Font | Arial throughout |
| Headings | Green `#009959`, 14/12/11/11 pt |
| Table Headers | Green fill, white bold text |
| Table Rows | Alternating white / `#E2EFD9` |
| Inner Borders | `#80CCB0` (NOT `#E2EFD9`) |
| Page Size | A4 Portrait, 2cm margins |
| Header | Right-aligned, grey text, green border |
| Footer | Copyright left, page number right (bold green) |

---

## 📋 Compliance Summary

After every run, the agent outputs a structured report:

```
COMPLIANCE SUMMARY
==================
Headings renumbered:      [list of changes]
Tables reformatted:       [count and names]
Figures reformatted:      [count and names]
Captions corrected:       [details]
Fonts corrected:          [non-Arial fonts replaced]
Colours corrected:        [non-compliant colours replaced]
Header/footer applied:    Yes / No
...
```

---

© 2026 INTEGRATION — Energy Plus Project
