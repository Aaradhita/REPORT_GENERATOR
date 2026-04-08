import os
from typing import Dict

from config import DOC_TEMPLATE_DIR, TEMPLATE_METADATA


class TemplateLoaderError(FileNotFoundError):
    """Raised when the requested report template cannot be found."""


def normalize_template_name(template_name: str) -> str:
    if not template_name:
        return ""
    normalized = template_name.strip().lower()
    for key in TEMPLATE_METADATA.keys():
        if key.lower() == normalized:
            return key
    return template_name.strip()


def load_template(template_name: str) -> str:
    template_key = normalize_template_name(template_name)
    template_data = TEMPLATE_METADATA.get(template_key)
    if not template_data:
        raise TemplateLoaderError(
            f"Unsupported report type '{template_name}'. Available templates: {', '.join(TEMPLATE_METADATA.keys())}"
        )

    template_path = os.path.join(DOC_TEMPLATE_DIR, template_data["filename"])
    if not os.path.exists(template_path):
        raise TemplateLoaderError(
            f"Template file not found for '{template_key}'. Expected path: {template_path}"
        )

    return template_path
