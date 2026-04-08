import os
import zipfile
from datetime import datetime
from typing import Iterable

from werkzeug.utils import secure_filename


def ensure_directory(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path


def get_unique_filename(base_name: str, extension: str = ".docx") -> str:
    safe_name = secure_filename(base_name) or "report"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if not extension.startswith("."):
        extension = f".{extension}"
    return f"{safe_name}_{timestamp}{extension}"


def is_allowed_file(filename: str, allowed_extensions: Iterable[str]) -> bool:
    return (
        bool(filename)
        and "." in filename
        and filename.rsplit(".", 1)[1].lower() in {ext.lower() for ext in allowed_extensions}
    )


def save_uploaded_file(uploaded_file, target_folder: str) -> str:
    filename = secure_filename(uploaded_file.filename)
    if not filename:
        raise ValueError("Invalid file name provided.")
    ensure_directory(target_folder)
    target_path = os.path.join(target_folder, filename)
    uploaded_file.save(target_path)
    return target_path


def extract_zip(zip_path: str, extract_to: str) -> None:
    ensure_directory(extract_to)
    with zipfile.ZipFile(zip_path, "r") as archive:
        archive.extractall(extract_to)
