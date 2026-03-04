# AGENT ZEE — AI Document Processing Assistant

**Multi-job document processing for international development consulting (WB, EU, ADB)**

A Python-based AI assistant that processes `.docx`, `.pdf`, and image files across specialized jobs — from formatting and analysis to proposal generation and CV rewriting.

---

## ⚡ Available Jobs

| # | Job | Description |
|---|-----|-------------|
| 1 | **Document Formatting** | Reformat uploaded files (`.docx`, `.pdf`, images) to INTEGRATION style guide with full compliance |
| 2 | **Proposal Development** | Generate structured proposals compliant with WB, EU, ADB requirements |
| 3 | **Documentation Analysis** | Readability metrics, structure breakdown, quality scoring |
| 4 | **Document Comparison** | Compare two documents with structural & content gap analysis |
| 5 | **Project Management** | Generate WBS, Gantt tables, RACI matrix, risk register |
| 6 | **Report Writing** | Transform raw content into professional WB/EU/ADB reports |
| 7 | **JIS Mapping** | Results Framework, LogFrame, and M&E mapping matrix |
| 8 | **CV Reception & Rewriting** | Reformat CVs to WB, EU, or ADB template requirements |
| 9 | **Contract Management** | Extract and structure contract deliverables, payments, timelines, and clauses |
| 10 | **GIS Coordinate Extraction** | Extract coordinates from content and export GIS-ready CSV |

---

## 🚀 Quick Start

### Install Dependencies

```bash
pip install -r requirements.txt
```

Optional OCR support for scanned PDFs/images requires the Tesseract binary:

```bash
# macOS
brew install tesseract

# Ubuntu/Debian
sudo apt-get install tesseract-ocr
```

### Web UI (Recommended)

```bash
python app.py
# Open http://localhost:5000 in your browser
```

### CLI Usage

```bash
# Document Formatting
python main.py format input.docx -o output.docx --report "Energy Master Plan" --year 2026

# Proposal Development
python main.py proposal draft.docx -o proposal.docx --client "World Bank (WB)" --title "Energy Project"

# Documentation Analysis
python main.py analyze report.docx -o analysis.docx --depth "Deep Dive"

# Document Comparison
python main.py compare doc_a.docx doc_b.docx -o comparison.docx

# Project Management
python main.py project project.docx -o pm_report.docx --name "GB Energy" --duration 24

# Report Writing
python main.py report raw.docx -o report.docx --style "World Bank" --title "Final Report"

# JIS Mapping
python main.py jis project.docx -o jis.docx --framework "Full Package" --sector "Energy"

# CV Rewriting
python main.py cv resume.docx -o cv_wb.docx --client "World Bank (WB)" --position "Team Leader"

# GIS Coordinate Extraction
python main.py gis survey.pdf -o points.csv
```

---

## 📁 Project Structure

```
integration-agent/
├── main.py              # CLI entry point (10 subcommands)
├── app.py               # Flask web UI (multi-job dashboard)
├── formatter/           # Document formatting engine
│   ├── constants.py     # Colour palette, typography, layout constants
│   ├── engine.py        # Core orchestrator (12-step pipeline)
│   ├── headings.py      # Heading detection & auto-numbering
│   ├── tables.py        # Table formatting (fills, borders, captions)
│   ├── figures.py       # Figure formatting (captions, source lines)
│   ├── typography.py    # Font enforcement (Arial everywhere)
│   ├── layout.py        # Page layout, header & footer
│   ├── lists.py         # Bullet list formatting
│   └── compliance.py    # Compliance summary generator
├── jobs/                # Job processing engines
│   ├── __init__.py      # Job registry & metadata
│   ├── proposal_development.py
│   ├── document_analysis.py
│   ├── comparison.py
│   ├── project_management.py
│   ├── report_writing.py
│   ├── jis_mapping.py
│   ├── cv_rewriting.py
│   ├── contract_management.py
│   └── gis_extraction.py
├── templates/
│   └── index.html       # Multi-job dashboard UI
├── requirements.txt
└── README.md
```

---

## 🎨 Client Templates Supported

**World Bank (WB)** — Standard proposal structure, WB CV format, Results Framework  
**European Union (EU)** — Europass CV, EU proposal format, CEFR language scales  
**Asian Development Bank (ADB)** — ADB proposal format, detailed assignment descriptions  

---

© 2026 AGENT ZEE — AI Document Assistant
