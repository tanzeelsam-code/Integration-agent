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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from formatter.engine import reformat_document
from jobs.proposal_development import process_proposal
from jobs.document_analysis import process_analysis
from jobs.comparison import process_comparison
from jobs.project_management import process_project_management
from jobs.report_writing import process_report
from jobs.jis_mapping import process_jis_mapping
from jobs.cv_rewriting import process_cv_rewrite


BANNER = r"""
  ╔══════════════════════════════════════════════════════════════╗
  ║             AGENT ZEE — AI Document Assistant               ║
  ║   Multi-Job CLI for WB, EU, and ADB Consulting              ║
  ╚══════════════════════════════════════════════════════════════╝
"""


def cmd_format(args):
    summary = reformat_document(
        input_path=args.input, output_path=args.output,
        report_name=args.report, year=args.year, project_number=args.project,
    )
    return summary


def cmd_proposal(args):
    return process_proposal(
        args.input, args.output,
        client_type=args.client, project_title=args.title, country=args.country,
    )


def cmd_analyze(args):
    return process_analysis(args.input, args.output, analysis_depth=args.depth)


def cmd_compare(args):
    return process_comparison(
        [args.input, args.input2], args.output,
        comparison_mode=args.mode,
    )


def cmd_project(args):
    return process_project_management(
        args.input, args.output,
        project_name=args.name, duration_months=args.duration,
    )


def cmd_report(args):
    return process_report(
        args.input, args.output,
        report_style=args.style, report_title=args.title,
        include_exec_summary=args.exec_summary,
    )


def cmd_jis(args):
    return process_jis_mapping(
        args.input, args.output,
        framework_type=args.framework, sector=args.sector,
    )


def cmd_cv(args):
    return process_cv_rewrite(
        args.input, args.output,
        client_template=args.client, position_title=args.position,
        years_experience_required=args.years,
    )


def main():
    parser = argparse.ArgumentParser(
        description="AGENT ZEE — AI Document Processing Assistant",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command", help="Available jobs")

    # Format
    p = sub.add_parser("format", help="Reformat .docx to INTEGRATION style")
    p.add_argument("input", help="Input .docx file")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--report", default="Report")
    p.add_argument("--year", default="2026")
    p.add_argument("--project", default="PRJ-001")

    # Proposal
    p = sub.add_parser("proposal", help="Generate structured proposal")
    p.add_argument("input", help="Input .docx (draft/TOR)")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--client", default="World Bank (WB)")
    p.add_argument("--title", default="")
    p.add_argument("--country", default="Pakistan")

    # Analyze
    p = sub.add_parser("analyze", help="Document analysis & quality report")
    p.add_argument("input", help="Input .docx")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--depth", default="Standard", choices=["Quick Overview","Standard","Deep Dive"])

    # Compare
    p = sub.add_parser("compare", help="Compare two documents")
    p.add_argument("input", help="First .docx")
    p.add_argument("input2", help="Second .docx")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--mode", default="Full Comparison")

    # Project
    p = sub.add_parser("project", help="Generate project management docs")
    p.add_argument("input", help="Input .docx")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--name", default="Untitled Project")
    p.add_argument("--duration", default="12")

    # Report
    p = sub.add_parser("report", help="Professional report writing")
    p.add_argument("input", help="Input .docx (raw content)")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--style", default="Generic Professional")
    p.add_argument("--title", default="Report")
    p.add_argument("--exec-summary", default="Yes", choices=["Yes","No"])

    # JIS
    p = sub.add_parser("jis", help="JIS mapping (LogFrame, Results Framework)")
    p.add_argument("input", help="Input .docx")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--framework", default="Full Package")
    p.add_argument("--sector", default="Energy")

    # CV
    p = sub.add_parser("cv", help="CV rewriting for WB/EU/ADB")
    p.add_argument("input", help="Input CV .docx")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--client", default="World Bank (WB)")
    p.add_argument("--position", default="")
    p.add_argument("--years", default="10")

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
        "jis": cmd_jis, "cv": cmd_cv,
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
