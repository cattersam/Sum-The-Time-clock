from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional
import logging
import re


SUPPORTED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
TIME_PATTERN = re.compile(r"^(?:[0-9]|[01][0-9]|2[0-3]):[0-5][0-9]$")
TIME_CANDIDATE_PATTERN = re.compile(r"(?:[0-9OolISB]{1,2})[:;：.,](?:[0-9OolISB]{2})")


def setup_logging(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"attendance_ocr_{datetime.now():%Y%m%d_%H%M%S}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(),
        ],
        force=True,
    )
    logging.info("log started")
    return log_path


def get_default_month() -> int:
    current_month = datetime.now().month
    return min(max(current_month, 1), 12)


def normalize_time_text(text: str) -> str:
    normalized = str(text or "").strip()
    replacements = {
        "O": "0",
        "o": "0",
        "I": "1",
        "l": "1",
        "|": "1",
        "S": "5",
        "B": "8",
        "：": ":",
        ";": ":",
        ".": ":",
        ",": ":",
        " ": "",
    }
    for old, new in replacements.items():
        normalized = normalized.replace(old, new)
    return normalized


def extract_time_text(text: str) -> Optional[str]:
    normalized = normalize_time_text(text)
    if TIME_PATTERN.fullmatch(normalized):
        hours, minutes = normalized.split(":")
        return f"{int(hours):02d}:{minutes}"
    candidate = TIME_CANDIDATE_PATTERN.search(str(text or ""))
    if candidate:
        return extract_time_text(candidate.group(0))
    return None


def extract_ocr_time_text(text: str) -> Optional[str]:
    normalized = normalize_time_text(text)
    valid = extract_time_text(normalized)
    if valid:
        return valid
    digits = re.sub(r"[^0-9]", "", normalized)
    if len(digits) == 3:
        candidate = f"{digits[0]}:{digits[1:]}"
    elif len(digits) == 4:
        candidate = f"{digits[:2]}:{digits[2:]}"
    else:
        return None
    return extract_time_text(candidate)


def is_valid_time(text: str) -> bool:
    return extract_time_text(text) is not None


def scan_image_files(folder: Path) -> List[Path]:
    if not folder.is_dir():
        return []
    return sorted(
        path
        for path in folder.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS
    )


def filter_image_files(paths: Iterable[str | Path]) -> List[Path]:
    return sorted(
        Path(path)
        for path in paths
        if Path(path).is_file() and Path(path).suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS
    )
