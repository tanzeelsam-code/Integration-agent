"""
Input normalization utilities.

All processing engines operate on .docx files. This module accepts supported
input formats and converts them into temporary .docx files when needed.
"""

from __future__ import annotations

import os
from pathlib import Path

from docx import Document
from docx.shared import Inches

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover - handled at runtime
    PdfReader = None


SUPPORTED_EXTENSIONS = {
    ".docx",
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".gif",
    ".tif",
    ".tiff",
    ".webp",
}


def is_supported_filename(filename: str) -> bool:
    return Path(filename or "").suffix.lower() in SUPPORTED_EXTENSIONS


def supported_extensions_csv() -> str:
    """Return a stable comma-separated accept string for HTML/file validators."""
    order = (
        ".docx",
        ".pdf",
        ".png",
        ".jpg",
        ".jpeg",
        ".bmp",
        ".gif",
        ".tif",
        ".tiff",
        ".webp",
    )
    return ",".join(order)


def normalize_input_to_docx(input_path: str, output_docx_path: str) -> dict:
    """
    Normalize supported file types into a .docx input.

    Returns:
        {
            "path": str,      # path to the .docx used for processing
            "converted": bool,
            "note": str | None,
        }
    """
    ext = Path(input_path).suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported file type '{ext}'. Supported: {supported_extensions_csv()}"
        )

    if ext == ".docx":
        return {"path": input_path, "converted": False, "note": None}

    if ext == ".pdf":
        note = _pdf_to_docx(input_path, output_docx_path)
        return {"path": output_docx_path, "converted": True, "note": note}

    note = _image_to_docx(input_path, output_docx_path)
    return {"path": output_docx_path, "converted": True, "note": note}


def _pdf_to_docx(input_path: str, output_docx_path: str) -> str:
    doc = Document()
    doc.add_heading("Imported PDF Content", level=1)
    doc.add_paragraph(f"Source file: {os.path.basename(input_path)}")
    doc.add_paragraph("")

    if PdfReader is None:
        doc.add_paragraph(
            "PDF text extraction is unavailable because the pypdf dependency is missing."
        )
        doc.save(output_docx_path)
        return "PDF uploaded, but text extraction was unavailable (missing pypdf)."

    reader = PdfReader(input_path)
    extracted_pages = 0

    for page_num, page in enumerate(reader.pages, 1):
        text = (page.extract_text() or "").strip()
        if not text:
            continue

        extracted_pages += 1
        if extracted_pages > 1:
            doc.add_page_break()
        doc.add_heading(f"Page {page_num}", level=2)
        for line in text.splitlines():
            clean = line.strip()
            if clean:
                doc.add_paragraph(clean)

    if extracted_pages == 0:
        doc.add_paragraph(
            "No extractable text was found in the PDF. "
            "If this is a scanned PDF, OCR is required for full text extraction."
        )
        note = "PDF converted with no extractable text (likely scanned image PDF)."
    else:
        note = f"PDF converted to DOCX using extracted text from {extracted_pages} page(s)."

    doc.save(output_docx_path)
    return note


def _image_to_docx(input_path: str, output_docx_path: str) -> str:
    doc = Document()
    doc.add_heading("Imported Image Content", level=1)
    doc.add_paragraph(f"Source file: {os.path.basename(input_path)}")
    doc.add_paragraph("")

    try:
        doc.add_picture(input_path, width=Inches(6.0))
    except Exception as exc:
        doc.add_paragraph(f"Unable to embed image preview: {exc}")

    doc.add_paragraph(
        "This file was converted from an image into DOCX for processing. "
        "OCR is not enabled, so text-based analysis may be limited."
    )
    doc.save(output_docx_path)
    return "Image converted to DOCX (embedded preview, no OCR)."
