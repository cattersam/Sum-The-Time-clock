from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Tuple
import json
import re


class ConfigError(RuntimeError):
    pass


class AppConfig:
    DEFAULT_VALIDATION = {
        "min_shift_minutes": 120,
        "max_shift_minutes": 960,
        "warn_short_shift_minutes": 240,
        "ok_confidence_threshold": 0.88,
    }

    DEFAULT_LEARNING = {
        "enabled": True,
        "history_file": "logs/correction_history.jsonl",
        "min_count": 2,
    }

    DEFAULT_UI = {
        "enable_drag_drop": True,
        "highlight_abnormal_rows": True,
    }

    def __init__(self, config_path: str | Path) -> None:
        self.path = Path(config_path)
        self.data = self._load()
        self._validate()

    def _load(self) -> dict:
        if not self.path.is_file():
            raise ConfigError(f"Missing config.json: {self.path}")
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise ConfigError(f"Could not read config.json: {exc}") from exc

    def _validate(self) -> None:
        for section in ("excel_mapping", "photo_mapping", "roi"):
            if section not in self.data:
                raise ConfigError(f"config.json is missing {section}")
        for half in ("first_half", "second_half"):
            roi = self.data["roi"].get(half)
            if not isinstance(roi, dict):
                raise ConfigError(f"config.json is missing roi.{half}")
            required = ("table_bbox", "start_col", "end_col", "row_count")
            if any(key not in roi for key in required):
                raise ConfigError(f"config.json roi.{half} is missing a required key")
            for key, expected_length in {"table_bbox": 4, "start_col": 2, "end_col": 2}.items():
                values = roi[key]
                if not isinstance(values, list) or len(values) != expected_length:
                    raise ConfigError(f"roi.{half}.{key} has an invalid shape")
                if not all(isinstance(value, (int, float)) for value in values):
                    raise ConfigError(f"roi.{half}.{key} must contain numbers")
                if not all(0 <= float(value) <= 1 for value in values):
                    raise ConfigError(f"roi.{half}.{key} values must be between 0 and 1")
            if int(roi["row_count"]) <= 0:
                raise ConfigError(f"roi.{half}.row_count must be greater than 0")

    @property
    def excel_mapping(self) -> Dict:
        return self.data["excel_mapping"]

    @property
    def validation(self) -> Dict:
        merged = dict(self.DEFAULT_VALIDATION)
        merged.update(self.data.get("validation") or {})
        return merged

    @property
    def learning(self) -> Dict:
        merged = dict(self.DEFAULT_LEARNING)
        merged.update(self.data.get("learning") or {})
        return merged

    @property
    def ui(self) -> Dict:
        merged = dict(self.DEFAULT_UI)
        merged.update(self.data.get("ui") or {})
        return merged

    def get_roi(self, half: str) -> Dict:
        if half not in {"first_half", "second_half"}:
            raise ConfigError(f"Unknown photo half: {half}")
        return self.data["roi"][half]

    def get_days(self, half: str) -> List[int]:
        start, end = self.data["photo_mapping"][half]["days"]
        return list(range(int(start), int(end) + 1))

    def classify_photo(self, filename: str | Path, fallback_index: int = 0) -> str:
        stem = Path(filename).stem.lower()
        first_hints = (r"1\s*[-_~到至]\s*15", "first", "front", "early", "上半", "前半")
        second_hints = (r"16\s*[-_~到至]\s*31", "second", "back", "late", "下半", "後半")
        if any(re.search(pattern, stem) for pattern in first_hints):
            return "first_half"
        if any(re.search(pattern, stem) for pattern in second_hints):
            return "second_half"
        return "first_half" if fallback_index == 0 else "second_half"

    def iter_row_boxes(self, half: str) -> List[Tuple[int, Tuple[float, float, float, float], Tuple[float, float, float, float]]]:
        roi = self.get_roi(half)
        days = self.get_days(half)
        row_count = int(roi["row_count"])
        if len(days) != row_count:
            raise ConfigError(f"{half} day range does not match roi row_count")
        start_x1, start_x2 = roi["start_col"]
        end_x1, end_x2 = roi["end_col"]
        row_height = 1.0 / row_count
        rows = []
        for index, day in enumerate(days):
            y1 = index * row_height
            y2 = (index + 1) * row_height
            rows.append((day, (start_x1, y1, start_x2, y2), (end_x1, y1, end_x2, y2)))
        return rows
