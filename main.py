#!/usr/bin/env python3
"""
AGENT ZEE — CLI Entry Point with multi-job subcommands.

Usage:
    python main.py format input.docx -o output.docx
    python main.py analyze input.docx -o output.docx
    python main.py proposal input.docx -o output.docx --client "World Bank (WB)"
    python main.py compare doc_a.docx doc_b.docx -o output.docx
    python main.py project input.docx -o output.docx --name "Energy Project"
    python main.py report input.docx -o output.docx --style "World Bank"
    python main.py jis input.docx -o output.docx --sector "Energy"
    python main.py cv input.docx -o output.docx --client "World Bank (WB)"
"""

import argparse
import sys
import os
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from formatter.engine import reformat_document
from jobs.proposal_development import process_proposal
from jobs.document_analysis import process_analysis
from jobs.comparison import process_comparison
from jobs.project_management import process_project_management
from jobs.report_writing import process_report
from jobs.jis_mapping import process_jis_mapping
from jobs.cv_rewriting import process_cv_rewrite
from jobs.contract_management import process_contract
from jobs.gis_extraction import process_gis
from input_adapter import (
    is_supported_filename,
    normalize_input_to_docx,
    supported_extensions_csv,
)


BANNER = r"""
  ╔══════════════════════════════════════════════════════════════╗
  ║             AGENT ZEE — AI Document Assistant               ║
  ║   Multi-Job CLI for WB, EU, and ADB Consulting              ║
  ╚══════════════════════════════════════════════════════════════╝
"""


def _prepare_input(input_path: str, temp_dir: str, slot: str) -> tuple[str, str | None]:
    if not is_supported_filename(input_path):
        raise ValueError(f"Unsupported input type. Supported: {supported_extensions_csv()}")

    converted_path = os.path.join(temp_dir, f"{slot}.docx")
    prepared = normalize_input_to_docx(input_path, converted_path)
    return prepared["path"], prepared.get("note")


def _append_conversion_notes(summary: str, notes: list[str | None]) -> str:
    filtered = [n for n in notes if n]
    if not filtered:
        return summary
    note_text = "\n".join(f"- {n}" for n in filtered)
    return f"{summary}\nINPUT CONVERSION NOTES\n{'=' * 40}\n{note_text}\n"


def cmd_format(args):
    with tempfile.TemporaryDirectory(prefix="integration_cli_") as temp_dir:
        input_path, note = _prepare_input(args.input, temp_dir, "input1")
        summary = reformat_document(
            input_path=input_path, output_path=args.output,
            report_name=args.report, year=args.year, project_number=args.project,
        )
        return _append_conversion_notes(summary, [note])


def cmd_proposal(args):
    with tempfile.TemporaryDirectory(prefix="integration_cli_") as temp_dir:
        input_path, note = _prepare_input(args.input, temp_dir, "input1")
        summary = process_proposal(
            input_path, args.output,
            client_type=args.client, project_title=args.title, country=args.country,
        )
        return _append_conversion_notes(summary, [note])


def cmd_analyze(args):
    with tempfile.TemporaryDirectory(prefix="integration_cli_") as temp_dir:
        input_path, note = _prepare_input(args.input, temp_dir, "input1")
        summary = process_analysis(input_path, args.output, analysis_depth=args.depth)
        return _append_conversion_notes(summary, [note])


def cmd_compare(args):
    with tempfile.TemporaryDirectory(prefix="integration_cli_") as temp_dir:
        input_path1, note1 = _prepare_input(args.input, temp_dir, "input1")
        input_path2, note2 = _prepare_input(args.input2, temp_dir, "input2")
        summary = process_comparison(
            [input_path1, input_path2], args.output,
            comparison_mode=args.mode,
        )
        return _append_conversion_notes(summary, [note1, note2])


def cmd_project(args):
    with tempfile.TemporaryDirectory(prefix="integration_cli_") as temp_dir:
        input_path, note = _prepare_input(args.input, temp_dir, "input1")
        summary = process_project_management(
            input_path, args.output,
            project_name=args.name, duration_months=args.duration,
        )
        return _append_conversion_notes(summary, [note])


def cmd_report(args):
    with tempfile.TemporaryDirectory(prefix="integration_cli_") as temp_dir:
        input_path, note = _prepare_input(args.input, temp_dir, "input1")
        summary = process_report(
            input_path, args.output,
            report_style=args.style, report_title=args.title,
            include_exec_summary=args.exec_summary,
        )
        return _append_conversion_notes(summary, [note])


def cmd_jis(args):
    with tempfile.TemporaryDirectory(prefix="integration_cli_") as temp_dir:
        input_path, note = _prepare_input(args.input, temp_dir, "input1")
        summary = process_jis_mapping(
            input_path, args.output,
            framework_type=args.framework, sector=args.sector,
        )
        return _append_conversion_notes(summary, [note])


def cmd_cv(args):
    with tempfile.TemporaryDirectory(prefix="integration_cli_") as temp_dir:
        input_path, note = _prepare_input(args.input, temp_dir, "input1")
        summary = process_cv_rewrite(
            input_path, args.output,
            client_template=args.client, position_title=args.position,
            years_experience_required=args.years,
        )
        return _append_conversion_notes(summary, [note])


def cmd_contract(args):
    with tempfile.TemporaryDirectory(prefix="integration_cli_") as temp_dir:
        input_path, note = _prepare_input(args.input, temp_dir, "input1")
        summary = process_contract(
            input_path, args.output,
            client_name=args.client_name, contract_type=args.type,
        )
        return _append_conversion_notes(summary, [note])


def cmd_gis(args):
    with tempfile.TemporaryDirectory(prefix="integration_cli_") as temp_dir:
        input_path, note = _prepare_input(args.input, temp_dir, "input1")
        summary = process_gis(input_path, args.output)
        return _append_conversion_notes(summary, [note])


def main():
    parser = argparse.ArgumentParser(
        description="AGENT ZEE — AI Document Processing Assistant",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command", help="Available jobs")
    input_help = "Input file (.docx, .pdf, or image)"

    # Format
    p = sub.add_parser("format", help="Reformat .docx to INTEGRATION style")
    p.add_argument("input", help=input_help)
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--report", default="Report")
    p.add_argument("--year", default="2026")
    p.add_argument("--project", default="PRJ-001")

    # Proposal
    p = sub.add_parser("proposal", help="Generate structured proposal")
    p.add_argument("input", help=input_help)
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--client", default="World Bank (WB)")
    p.add_argument("--title", default="")
    p.add_argument("--country", default="Pakistan")

    # Analyze
    p = sub.add_parser("analyze", help="Document analysis & quality report")
    p.add_argument("input", help=input_help)
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--depth", default="Standard", choices=["Quick Overview","Standard","Deep Dive"])

    # Compare
    p = sub.add_parser("compare", help="Compare two documents")
    p.add_argument("input", help="First input file (.docx, .pdf, or image)")
    p.add_argument("input2", help="Second input file (.docx, .pdf, or image)")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--mode", default="Full Comparison")

    # Project
    p = sub.add_parser("project", help="Generate project management docs")
    p.add_argument("input", help=input_help)
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--name", default="Untitled Project")
    p.add_argument("--duration", default="12")

    # Report
    p = sub.add_parser("report", help="Professional report writing")
    p.add_argument("input", help=input_help)
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--style", default="Generic Professional")
    p.add_argument("--title", default="Report")
    p.add_argument("--exec-summary", default="Yes", choices=["Yes","No"])

    # JIS
    p = sub.add_parser("jis", help="JIS mapping (LogFrame, Results Framework)")
    p.add_argument("input", help=input_help)
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--framework", default="Full Package")
    p.add_argument("--sector", default="Energy")

    # CV
    p = sub.add_parser("cv", help="CV rewriting for WB/EU/ADB")
    p.add_argument("input", help=input_help)
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--client", default="World Bank (WB)")
    p.add_argument("--position", default="")
    p.add_argument("--years", default="10")

    # Contract
    p = sub.add_parser("contract", help="Extract and manage contract details")
    p.add_argument("input", help=input_help)
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--client-name", default="Unknown Client")
    p.add_argument("--type", default="Standard Consulting")

    # GIS Extraction
    p = sub.add_parser("gis", help="Extract GPS coordinate points to CSV")
    p.add_argument("input", help=input_help)
    p.add_argument("-o", "--output", required=True)

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(0)

    print(BANNER)
    print(f"  ⚙️  Job: {args.command}")
    print(f"  📄 Input: {args.input}")
    print(f"  📁 Output: {args.output}")
    print()
    print("  ⏳ Processing...")
    print()

    handlers = {
        "format": cmd_format, "proposal": cmd_proposal, "analyze": cmd_analyze,
        "compare": cmd_compare, "project": cmd_project, "report": cmd_report,
        "jis": cmd_jis, "cv": cmd_cv, "contract": cmd_contract, "gis": cmd_gis,
    }

    try:
        summary = handlers[args.command](args)
        print("  ✅ Complete!")
        print(f"  📄 Saved to: {args.output}")
        print()
        print("─" * 60)
        print(summary)
        print("─" * 60)
    except Exception as e:
        print(f"  ❌ Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
