#!/usr/bin/env python3
"""
Formating AI Assistance — CLI Entry Point

Usage:
    python main.py input.docx -o output.docx
    python main.py input.docx -o output.docx --report "Energy Master Plan" --year 2026 --project "EMP-GB-001"
"""

import argparse
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from formatter.engine import reformat_document


BANNER = r"""
  ╔══════════════════════════════════════════════════════════════╗
  ║   INTEGRATION Energy Plus — Formating AI Assistance         ║
  ║   Energy Master Plan for Gilgit-Baltistan, Pakistan         ║
  ╚══════════════════════════════════════════════════════════════╝
"""


def main():
    parser = argparse.ArgumentParser(
        description="Formating AI Assistance — "
                    "Reformat .docx files to INTEGRATION Energy Plus style guide.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Example:\n  python main.py report.docx -o report_formatted.docx "
               '--report "Energy Master Plan" --year 2026 --project "EMP-GB-001"',
    )
    parser.add_argument("input", help="Path to input .docx file")
    parser.add_argument("-o", "--output", required=True, help="Path to save reformatted .docx")
    parser.add_argument("--report", default="Report", help="Report name for header (default: Report)")
    parser.add_argument("--year", default="2026", help="Year for footer copyright (default: 2026)")
    parser.add_argument("--project", default="PRJ-001", help="Project number for footer (default: PRJ-001)")

    args = parser.parse_args()

    # Validate input
    if not os.path.isfile(args.input):
        print(f"❌  Error: Input file not found: {args.input}")
        sys.exit(1)

    if not args.input.lower().endswith(".docx"):
        print(f"⚠️  Warning: Input file does not have .docx extension: {args.input}")

    print(BANNER)
    print(f"  📄 Input:   {args.input}")
    print(f"  📁 Output:  {args.output}")
    print(f"  📋 Report:  {args.report}")
    print(f"  📅 Year:    {args.year}")
    print(f"  🔢 Project: {args.project}")
    print()
    print("  ⏳ Processing...")
    print()

    try:
        summary = reformat_document(
            input_path=args.input,
            output_path=args.output,
            report_name=args.report,
            year=args.year,
            project_number=args.project,
        )
        print("  ✅ Reformatting complete!")
        print(f"  📄 Saved to: {args.output}")
        print()
        print("─" * 60)
        print(summary)
        print("─" * 60)
    except Exception as e:
        print(f"  ❌ Error during reformatting: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
