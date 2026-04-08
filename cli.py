import argparse
import json
import os
import sys
from typing import Optional

from config import TEMPLATE_METADATA
from services.report_generator import generate_report, normalize_report_type, ReportGenerationError
from utils.file_handler import ensure_directory, extract_zip


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate VAPT DOCX reports from JSON or Excel data."
    )
    parser.add_argument("--type", required=True, help="Report type (web, network, phishing, cloud, etc.)")
    parser.add_argument("--input", required=True, help="Path to a JSON or XLSX input file.")
    parser.add_argument("--output-name", help="Optional base name for the generated report.")
    parser.add_argument("--export-pdf", action="store_true", help="Also export a PDF version of the report.")
    parser.add_argument("--zip", dest="zip_path", help="Optional ZIP archive containing screenshots for screenshot-enabled templates.")
    return parser.parse_args()


def main() -> int:
    args = parse_arguments()
    template_type = normalize_report_type(args.type)
    if template_type not in TEMPLATE_METADATA:
        print(
            f"Unsupported template type '{args.type}'. Available templates: {', '.join(TEMPLATE_METADATA.keys())}",
            file=sys.stderr,
        )
        return 1

    if not os.path.exists(args.input):
        print(f"Input file not found: {args.input}", file=sys.stderr)
        return 1

    screenshot_dir: Optional[str] = None
    if args.zip_path:
        if not os.path.exists(args.zip_path):
            print(f"ZIP file not found: {args.zip_path}", file=sys.stderr)
            return 1
        screenshot_dir = os.path.join(os.getcwd(), "cli_screenshots")
        extract_zip(args.zip_path, screenshot_dir)

    output_dir = ensure_directory(os.path.join(os.getcwd(), "generated_reports"))

    try:
        result = generate_report(
            args.input,
            template_type,
            output_name=args.output_name,
            output_dir=output_dir,
            export_pdf_flag=args.export_pdf,
            screenshot_dir=screenshot_dir,
        )
        print(f"Report generated: {result['docx_path']}")
        if result.get("pdf_path"):
            print(f"PDF exported: {result['pdf_path']}")
        return 0
    except ReportGenerationError as exc:
        print(f"Report generation failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
