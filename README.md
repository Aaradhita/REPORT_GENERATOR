# REPORT_GENERATOR
This project is a web-based report generation tool built using Flask for the cybersecurity domain. It is designed to assist in creating structured and professional reports for Vulnerability Assessment and Penetration Testing (VAPT).
# VAPT Report Generator

A modular Flask-based Vulnerability Assessment and Penetration Testing (VAPT) report generator that builds DOCX reports from template documents and structured data.

## Features

- Clean, modular architecture
- Centralized application configuration in `config.py`
- Dynamic `.docx` template loading from `doc_templates/`
- Data-driven report generation for multiple report types
- Excel and JSON input support
- Optional screenshot embedding for screenshot-enabled templates
- PDF export support via `docx2pdf`
- CLI support via `cli.py`
- Friendly Flask UI for authenticated report generation and downloads

## Folder Structure

- `app.py` — Flask application entrypoint
- `cli.py` — Command-line report generator
- `config.py` — Central settings and template metadata
- `doc_templates/` — All `.docx` report templates
- `generated_reports/` — Generated output files
- `services/` — Report generation and template loading logic
- `utils/` — File handling helpers and utilities
- `templates/` — Flask HTML templates
- `static/` — Static site assets

## Installation

1. Create a Python virtual environment:

```bash
python -m venv .venv
```

2. Activate the environment:

- Windows:
  ```powershell
  .\.venv\Scripts\Activate.ps1
  ```
- macOS / Linux:
  ```bash
  source .venv/bin/activate
  ```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

## Running the Web UI

Start the Flask app:

```bash
python app.py
```

Then open:

```text
http://localhost:5555
```

Use the login credentials:

- Username: `sudo`
- Password: `technical`

## CLI Usage

Generate a report from JSON:

```bash
python cli.py --type web --input data.json --output-name my_web_report
```

Generate a report from Excel and optional screenshot ZIP:

```bash
python cli.py --type network --input data.xlsx --zip screenshots.zip --export-pdf
```

## Notes

- Templates must exist in `doc_templates/` and follow the expected placeholder structure.
- PDF export requires `docx2pdf` and a supported environment such as Windows with Microsoft Word installed.
- Reports are saved to `generated_reports/` with a timestamped filename.
