import json
import logging
import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Union

import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

from config import REPORT_FOLDER, TEMPLATE_METADATA
from services.template_loader import load_template
from utils.file_handler import ensure_directory, get_unique_filename

try:
    from docx2pdf import convert as convert_docx_to_pdf

    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

logger = logging.getLogger(__name__)


class ReportGenerationError(Exception):
    pass


def normalize_report_type(report_type: str) -> str:
    if not report_type:
        return ""
    normalized = report_type.strip().lower()
    for key in TEMPLATE_METADATA.keys():
        if key.lower() == normalized:
            return key
    return report_type.strip()


def _clean_value(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _ensure_required_columns(df: pd.DataFrame, columns_map: Dict[str, str]) -> List[str]:
    missing = []
    for col in columns_map.values():
        if col and col not in df.columns:
            missing.append(col)
    return missing


def _build_row_value(row: pd.Series, key: str, columns_map: Dict[str, str]) -> str:
    source_column = columns_map.get(key)
    if not source_column:
        return ""
    return _clean_value(row.get(source_column, ""))


def _build_screenshots(
    row: pd.Series,
    doc: DocxTemplate,
    columns_map: Dict[str, str],
    extract_path: Optional[str],
) -> List[Dict[str, Any]]:
    screenshots: List[Dict[str, Any]] = []
    if not extract_path:
        return screenshots

    poc_folder = _build_row_value(row, "poc_folder", columns_map)
    if not poc_folder:
        return screenshots

    folder_path = os.path.join(extract_path, poc_folder)
    if not os.path.isdir(folder_path):
        return screenshots

    for index, file_name in enumerate(sorted(os.listdir(folder_path)), start=1):
        if file_name.lower().endswith((".png", ".jpg", ".jpeg")):
            image_path = os.path.join(folder_path, file_name)
            try:
                screenshots.append(
                    {
                        "label": f"Screenshot {index}:",
                        "image": InlineImage(doc, image_path, width=Mm(155)),
                    }
                )
            except Exception as exc:
                logger.warning("Skipping screenshot %s: %s", image_path, exc)

    return screenshots


def _build_standard_context(
    df: pd.DataFrame,
    report_type: str,
    doc: DocxTemplate,
    extract_path: Optional[str],
) -> Dict[str, Any]:
    columns_map = TEMPLATE_METADATA[report_type]["columns"]
    if missing := _ensure_required_columns(df, columns_map):
        raise ReportGenerationError(
            f"Missing required columns for {report_type}: {', '.join(missing)}"
        )

    summary = []
    details = []
    uses_screenshots = TEMPLATE_METADATA[report_type]["uses_screenshots"]

    for index, row in df.iterrows():
        summary.append(
            {
                "sno": _build_row_value(row, "sno", columns_map) or str(index + 1),
                "asset": _build_row_value(row, "asset", columns_map),
                "title": _build_row_value(row, "title", columns_map),
                "cve": _build_row_value(row, "cve", columns_map),
                "severity": _build_row_value(row, "severity", columns_map),
                "reference": _build_row_value(row, "reference", columns_map),
                "new_or_repeat": _build_row_value(row, "status", columns_map),
                "affected_assets": _build_row_value(row, "affected_assets", columns_map),
                "vulnerable_url": _build_row_value(row, "vulnerable_url", columns_map),
                "vulnerable_function": _build_row_value(row, "vulnerable_function", columns_map),
                "vulnerable_component": _build_row_value(row, "vulnerable_component", columns_map),
                "recommendations": _build_row_value(row, "recommendation", columns_map)
                or _build_row_value(row, "recommendations", columns_map),
            }
        )

        details.append(
            {
                "title": _build_row_value(row, "title", columns_map),
                "asset": _build_row_value(row, "asset", columns_map),
                "affected_assets": _build_row_value(row, "affected_assets", columns_map),
                "vulnerable_url": _build_row_value(row, "vulnerable_url", columns_map),
                "vulnerable_function": _build_row_value(row, "vulnerable_function", columns_map),
                "vulnerable_component": _build_row_value(row, "vulnerable_component", columns_map),
                "description": _build_row_value(row, "description", columns_map),
                "cve": _build_row_value(row, "cve", columns_map),
                "severity": _build_row_value(row, "severity", columns_map),
                "recommendation": _build_row_value(row, "recommendation", columns_map)
                or _build_row_value(row, "recommendations", columns_map),
                "reference": _build_row_value(row, "reference", columns_map),
                "status": _build_row_value(row, "status", columns_map),
                "screenshots": _build_screenshots(
                    row, doc, columns_map, extract_path if uses_screenshots else None
                ),
            }
        )

    return {
        "report_release_date": datetime.now().strftime("%d/%m/%Y"),
        "vulnerabilities_summary": summary,
        "vulnerabilities_details": details,
    }


def _build_server_hardening_context(df: pd.DataFrame) -> Dict[str, Any]:
    columns_map = TEMPLATE_METADATA["SERVER HARDENING"]["columns"]
    if missing := _ensure_required_columns(df, columns_map):
        raise ReportGenerationError(
            f"Missing required columns for SERVER HARDENING: {', '.join(missing)}"
        )

    server_rows = [
        {key: _build_row_value(row, key, columns_map) for key in columns_map.keys()}
        for _, row in df.iterrows()
    ]

    return {
        "report_release_date": datetime.now().strftime("%d/%m/%Y"),
        "server_rows": server_rows,
    }


def _build_phishing_context(
    df: pd.DataFrame,
    doc: DocxTemplate,
    export_directory: str,
) -> Dict[str, Any]:
    columns_map = TEMPLATE_METADATA["PHISHING"]["columns"]
    if missing := _ensure_required_columns(df, columns_map):
        raise ReportGenerationError(
            f"Missing required columns for PHISHING: {', '.join(missing)}"
        )

    def get(row: pd.Series, key: str) -> str:
        return _clean_value(row.get(columns_map[key], ""))

    def normalize_status(value: str) -> str:
        value = value.strip().lower()
        if value in {"email sent", "sent"}:
            return "sent"
        if value in {"email opened", "opened"}:
            return "opened"
        if value in {"email clicked", "clicked"}:
            return "clicked"
        if value in {"email submitted", "submitted"}:
            return "submitted"
        return value

    def to_bool(value: Any) -> bool:
        return str(value).strip().lower() in {"true", "yes", "1", "y"}

    rows = []
    for _, row in df.iterrows():
        rows.append(
            {
                "status": normalize_status(get(row, "status")),
                "email": get(row, "email"),
                "group": get(row, "internal"),
                "reported": to_bool(row.get(columns_map["reported"], "")),
            }
        )

    internal_rows = [row for row in rows if row["group"].lower() == "internal"]
    external_rows = [row for row in rows if row["group"].lower() == "external"]

    def count_status(rows_group: List[Dict[str, Any]], statuses: Union[str, List[str]]) -> int:
        if isinstance(statuses, str):
            statuses = [statuses]
        return sum(1 for row in rows_group if row["status"] in statuses)

    def create_breakdown(rows_group: List[Dict[str, Any]], status: str) -> List[Dict[str, str]]:
        return [
            {"email": row["email"], "reported": "Yes" if row["reported"] else "No"}
            for row in rows_group
            if row["status"] == status
        ]

    internal_values = {
        "sent": count_status(internal_rows, ["sent", "opened", "clicked", "submitted"]),
        "opened": count_status(internal_rows, ["opened", "clicked", "submitted"]),
        "clicked": count_status(internal_rows, ["clicked", "submitted"]),
        "submitted": count_status(internal_rows, "submitted"),
    }
    external_values = {
        "sent": count_status(external_rows, ["sent", "opened", "clicked", "submitted"]),
        "opened": count_status(external_rows, ["opened", "clicked", "submitted"]),
        "clicked": count_status(external_rows, ["clicked", "submitted"]),
        "submitted": count_status(external_rows, "submitted"),
    }

    def safe_pct(value: int, total: int) -> float:
        return round(100 * value / total, 1) if total else 0.0

    chart_categories = [
        ("EMAIL SENT", internal_values["sent"], external_values["sent"]),
        ("OPENED EMAIL", internal_values["opened"], external_values["opened"]),
        ("CLICKED LINK", internal_values["clicked"], external_values["clicked"]),
        ("SUBMIT DATA", internal_values["submitted"], external_values["submitted"]),
    ]

    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        labels = [row[0] for row in chart_categories]
        internal_counts = [row[1] for row in chart_categories]
        external_counts = [row[2] for row in chart_categories]

        x_positions = range(len(labels))
        plt.figure(figsize=(9, 4.5))
        plt.bar([x - 0.2 for x in x_positions], internal_counts, width=0.4, label="Internal")
        plt.bar([x + 0.2 for x in x_positions], external_counts, width=0.4, label="External")
        plt.xticks(list(x_positions), labels, rotation=0)
        plt.ylabel("Count")
        plt.title("Phishing Simulation Report")
        plt.legend()
        plt.tight_layout()

        chart_path = os.path.join(export_directory, "phishing_chart.png")
        plt.savefig(chart_path, dpi=200)
        plt.close()
        chart_image = InlineImage(doc, chart_path, width=Mm(150))
    except Exception as exc:
        logger.warning("Unable to generate phishing chart image: %s", exc)
        chart_image = None

    return {
        "report_release_date": datetime.now().strftime("%d/%m/%Y"),
        "rows": rows,
        "ph_totals": {
            "sent": sum(count_status(rows, ["sent", "opened", "clicked", "submitted"])),
            "opened": sum(count_status(rows, ["opened", "clicked", "submitted"])),
            "clicked": sum(count_status(rows, ["clicked", "submitted"])),
            "submitted": count_status(rows, "submitted"),
        },
        "internal": {
            "sent": create_breakdown(internal_rows, "sent"),
            "opened": create_breakdown(internal_rows, "opened"),
            "clicked": create_breakdown(internal_rows, "clicked"),
            "submitted": create_breakdown(internal_rows, "submitted"),
        },
        "external": {
            "sent": create_breakdown(external_rows, "sent"),
            "opened": create_breakdown(external_rows, "opened"),
            "clicked": create_breakdown(external_rows, "clicked"),
            "submitted": create_breakdown(external_rows, "submitted"),
        },
        "internal_totals": {
            "total_emails": len({row["email"] for row in internal_rows if row["email"]}),
            **internal_values,
            "sent_pct": safe_pct(internal_values["sent"], internal_values["sent"]),
            "opened_pct": safe_pct(internal_values["opened"], internal_values["sent"]),
            "clicked_pct": safe_pct(internal_values["clicked"], internal_values["sent"]),
            "submitted_pct": safe_pct(internal_values["submitted"], internal_values["sent"]),
        },
        "external_totals": {
            "total_emails": len({row["email"] for row in external_rows if row["email"]}),
            **external_values,
            "sent_pct": safe_pct(external_values["sent"], external_values["sent"]),
            "opened_pct": safe_pct(external_values["opened"], external_values["sent"]),
            "clicked_pct": safe_pct(external_values["clicked"], external_values["sent"]),
            "submitted_pct": safe_pct(external_values["submitted"], external_values["sent"]),
        },
        "chart_data": [[label, internal_count, external_count] for label, internal_count, external_count in chart_categories],
        "phishing_chart": chart_image,
    }


def _build_dynamic_context(data: Dict[str, Any]) -> Dict[str, Any]:
    context = data.get("context", {}) if isinstance(data, dict) else {}
    if isinstance(data, dict) and "rows" in data:
        context["rows"] = data["rows"]
    context.setdefault("report_release_date", datetime.now().strftime("%d/%m/%Y"))
    return context


def export_pdf(docx_path: str, pdf_path: Optional[str] = None) -> str:
    if not PDF_SUPPORT:
        raise ReportGenerationError(
            "PDF export is not available. Install docx2pdf and ensure Word is available on Windows."
        )

    if not os.path.exists(docx_path):
        raise ReportGenerationError(f"Source DOCX not found for PDF export: {docx_path}")

    pdf_path = pdf_path or f"{os.path.splitext(docx_path)[0]}.pdf"
    try:
        convert_docx_to_pdf(docx_path, pdf_path)
    except Exception as exc:
        raise ReportGenerationError(f"PDF export failed: {exc}") from exc

    return pdf_path


def generate_report(
    data: Union[str, pd.DataFrame, Dict[str, Any]],
    template_type: str,
    output_name: Optional[str] = None,
    output_dir: Optional[str] = None,
    export_pdf_flag: bool = False,
    screenshot_dir: Optional[str] = None,
) -> Dict[str, Optional[str]]:
    report_type = normalize_report_type(template_type)
    template_path = load_template(report_type)
    output_dir = ensure_directory(output_dir or REPORT_FOLDER)
    output_name = output_name or report_type.replace(" ", "_")
    output_path = os.path.join(output_dir, get_unique_filename(output_name, ".docx"))

    doc = DocxTemplate(template_path)

    if isinstance(data, str):
        if data.lower().endswith(".xlsx"):
            try:
                df = pd.read_excel(data)
                if report_type == "SERVER HARDENING":
                    context = _build_server_hardening_context(df)
                elif report_type == "PHISHING":
                    context = _build_phishing_context(df, doc, output_dir)
                else:
                    context = _build_standard_context(
                        df, report_type, doc, screenshot_dir
                    )
            except Exception as exc:
                raise ReportGenerationError(f"Failed to load Excel input: {exc}") from exc
        elif data.lower().endswith(".json"):
            try:
                with open(data, "r", encoding="utf-8") as handle:
                    json_data = json.load(handle)
                context = _build_dynamic_context(json_data)
            except Exception as exc:
                raise ReportGenerationError(f"Failed to load JSON input: {exc}") from exc
        else:
            raise ReportGenerationError("Data must be a JSON or XLSX file path, pandas DataFrame, or a context dictionary.")
    elif isinstance(data, pd.DataFrame):
        if report_type == "SERVER HARDENING":
            context = _build_server_hardening_context(data)
        elif report_type == "PHISHING":
            context = _build_phishing_context(data, doc, output_dir)
        else:
            context = _build_standard_context(data, report_type, doc, screenshot_dir)
    elif isinstance(data, dict):
        context = _build_dynamic_context(data)
    else:
        raise ReportGenerationError("Unsupported data type for report generation.")

    try:
        doc.render(context)
        doc.save(output_path)
    except Exception as exc:
        logger.exception("Report render failed for %s", output_path)
        raise ReportGenerationError(f"Failed to save report: {exc}") from exc

    result = {"docx_path": output_path, "pdf_path": None}
    logger.info("Generated DOCX report: %s", output_path)

    if export_pdf_flag:
        result["pdf_path"] = export_pdf(output_path)
        logger.info("Generated PDF report: %s", result["pdf_path"])

    return result
