import logging
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOC_TEMPLATE_DIR = os.path.join(BASE_DIR, "doc_templates")
UPLOAD_FOLDER = os.path.join(BASE_DIR, "upload_folder")
REPORT_FOLDER = os.path.join(BASE_DIR, "generated_reports")
LOG_FILE = os.path.join(BASE_DIR, "app.log")

ALLOWED_EXTENSIONS = {"xlsx", "zip", "json"}
SECRET_KEY = os.environ.get("SECRET_KEY", "change-this-secret")
SESSION_PERMANENT = False
REMEMBER_COOKIE_DURATION = 0

TEMPLATE_METADATA = {
    "NETWORK": {
        "filename": "network_template.docx",
        "uses_screenshots": True,
        "columns": {
            "sno": "Sno",
            "asset": "Affected Asset",
            "title": "Observation / Vulnerability Title",
            "affected_assets": "Affected Assets",
            "description": "Detailed observation / Vulnerable point",
            "cve": "CVE/CWE",
            "severity": "Severity",
            "recommendation": "Recommendations",
            "reference": "Reference",
            "status": "New or Repeat Observation",
            "poc_folder": "POC Folder",
        },
    },
    "Console": {
        "filename": "Console_template.docx",
        "uses_screenshots": True,
        "columns": {
            "sno": "Sno",
            "asset": "Affected Asset",
            "title": "Observation / Vulnerability Title",
            "affected_assets": "Affected Assets",
            "description": "Detailed observation / Vulnerable point",
            "cve": "CVE/CWE",
            "severity": "Severity",
            "recommendation": "Recommendations",
            "reference": "Reference",
            "status": "New or Repeat Observation",
            "poc_folder": "POC Folder",
        },
    },
    "FIREWALL": {
        "filename": "firewall_template.docx",
        "uses_screenshots": True,
        "columns": {
            "sno": "Sno",
            "asset": "Affected Asset",
            "title": "Observation / Vulnerability Title",
            "affected_assets": "Affected Assets",
            "description": "Detailed observation / Vulnerable point",
            "cve": "CVE/CWE",
            "severity": "Severity",
            "recommendation": "Recommendations",
            "reference": "Reference",
            "status": "New or Repeat Observation",
            "poc_folder": "POC Folder",
        },
    },
    "WEB": {
        "filename": "web_template.docx",
        "uses_screenshots": True,
        "columns": {
            "sno": "Sno",
            "asset": "Affected Asset",
            "title": "Observation / Vulnerability Title",
            "vulnerable_url": "Vulnerable URL",
            "description": "Detailed observation / Vulnerable point",
            "cve": "CVE/CWE",
            "severity": "Severity",
            "recommendation": "Recommendations",
            "reference": "Reference",
            "status": "New or Repeat Observation",
            "poc_folder": "POC Folder",
        },
    },
    "SERVER HARDENING": {
        "filename": "server hardening_template.docx",
        "uses_screenshots": False,
        "columns": {
            "vuln_id": "VULNERABILITY_ID",
            "description": "DESCRIPTION",
            "category": "CATEGORY",
            "advices": "ADVICE",
            "solution": "SOLUTION",
            "name": "NAME",
            "severity": "SEVERITY",
            "host_name": "HOST NAME",
        },
    },
    "CLOUD": {
        "filename": "cloud_template.docx",
        "uses_screenshots": False,
        "columns": {
            "sno": "Sno",
            "asset": "Affected Asset",
            "title": "Observation / Vulnerability Title",
            "vulnerable_function": "Vulnerable Function",
            "vulnerable_component": "Vulnerable Component",
            "description": "Detailed observation / Vulnerable point",
            "cve": "CVE/CWE",
            "severity": "Severity",
            "recommendation": "Recommendations",
            "reference": "Reference",
            "status": "New or Repeat Observation",
            "poc_folder": "POC Folder",
        },
    },
    "OG CERTAIN - WEB": {
        "filename": "ogcertain-web_template.docx",
        "uses_screenshots": True,
        "columns": {
            "sno": "Sno",
            "asset": "Affected Asset",
            "title": "Observation / Vulnerability Title",
            "vulnerable_url": "Vulnerable URL",
            "description": "Detailed observation / Vulnerable point",
            "cve": "CVE/CWE",
            "severity": "Severity",
            "recommendation": "Recommendation",
            "recommendations": "Recommendations",
            "reference": "Reference",
            "status": "New or Repeat Observation",
            "poc_folder": "POC Folder",
        },
    },
    "OG CERTAIN - NETWORK": {
        "filename": "ogcertain-network_template.docx",
        "uses_screenshots": True,
        "columns": {
            "sno": "Sno",
            "asset": "Affected Asset",
            "title": "Observation / Vulnerability Title",
            "affected_assets": "Affected Assets",
            "description": "Detailed observation / Vulnerable point",
            "cve": "CVE/CWE",
            "severity": "Severity",
            "recommendation": "Recommendations",
            "recommendations": "Recommendation",
            "reference": "Reference",
            "status": "New or Repeat Observation",
            "poc_folder": "POC Folder",
        },
    },
    "PHISHING": {
        "filename": "phishing_template.docx",
        "uses_screenshots": False,
        "columns": {
            "status": "Status",
            "email": "Email",
            "internal": "Internal/External",
            "reported": "Reported",
        },
    },
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
