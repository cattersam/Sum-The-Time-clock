"""JSONL correction history for attendance OCR."""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable
import json
import re


TIME_RE = re.compile(r"^(?:[01]?\d|2[0-3]):[0-5]\d$")


@dataclass
class CorrectionEvent:
    timestamp: str
    date: str
    field: str
    ocr_text: str
    old_value: str
    new_value: str
    image_path: str
    confidence: float = 0.0
    reason: str = "manual_edit"


class CorrectionHistory:
    def __init__(self, path: str | Path, min_count: int = 2) -> None:
        self.path = Path(path)
        self.min_count = max(1, int(min_count))
        self.path.parent.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def normalize_raw(text: str) -> str:
        text = str(text or "").strip()
        text = text.translate(
            str.maketrans(
                {
                    "O": "0",
                    "o": "0",
                    "I": "1",
                    "l": "1",
                    "|": "1",
                    "S": "5",
                    "B": "8",
                    ";": ":",
                    "：": ":",
                    "；": ":",
                    ".": ":",
                    ",": ":",
                    "，": ":",
                    "。": ":",
                }
            )
        )
        return re.sub(r"\s+", "", text)

    @staticmethod
    def looks_like_time(text: str) -> bool:
        return bool(TIME_RE.match(str(text or "").strip()))

    @staticmethod
    def normalize_field(field: str) -> str:
        if field in {"clock_in", "start", "start_time", "上班", "出勤"}:
            return "clock_in"
        return "clock_out"

    def append(
        self,
        *,
        date: str | int,
        field: str,
        ocr_text: str,
        old_value: str,
        new_value: str,
        image_path: str,
        confidence: float = 0.0,
        reason: str = "manual_edit",
    ) -> bool:
        new_value = str(new_value or "").strip()
        if not self.looks_like_time(new_value):
            return False

        event = CorrectionEvent(
            timestamp=datetime.now().isoformat(timespec="seconds"),
            date=str(date),
            field=self.normalize_field(field),
            ocr_text=str(ocr_text or ""),
            old_value=str(old_value or ""),
            new_value=new_value,
            image_path=str(image_path or ""),
            confidence=float(confidence or 0.0),
            reason=reason,
        )
        with self.path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(asdict(event), ensure_ascii=False) + "\n")
        return True

    def iter_events(self) -> Iterable[Dict[str, Any]]:
        if not self.path.exists():
            return []
        events: list[Dict[str, Any]] = []
        with self.path.open("r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    events.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
        return events

    def build_rules(self) -> Dict[tuple[str, str], str]:
        counts: dict[tuple[str, str], Counter[str]] = defaultdict(Counter)
        for event in self.iter_events():
            raw = self.normalize_raw(event.get("ocr_text", ""))
            field = self.normalize_field(str(event.get("field", "")))
            corrected = str(event.get("new_value", "")).strip()
            if raw and self.looks_like_time(corrected):
                counts[(field, raw)][corrected] += 1

        rules: Dict[tuple[str, str], str] = {}
        for key, counter in counts.items():
            ranked = counter.most_common(2)
            if not ranked:
                continue
            best_value, best_count = ranked[0]
            second_count = ranked[1][1] if len(ranked) > 1 else 0
            if best_count >= self.min_count and best_count > second_count:
                rules[key] = best_value
        return rules

    def apply_to_record(self, record: Any) -> bool:
        rules = self.build_rules()
        changed = False
        learned_fields = set(getattr(record, "learned_fields", set()) or set())
        for field, raw_attr, value_attr in (
            ("clock_in", "start_raw", "start_time"),
            ("clock_out", "end_raw", "end_time"),
        ):
            raw_text = getattr(record, raw_attr, "") or getattr(record, value_attr, "")
            key = (field, self.normalize_raw(raw_text))
            corrected = rules.get(key)
            if corrected and corrected != getattr(record, value_attr, ""):
                setattr(record, value_attr, corrected)
                learned_fields.add(field)
                changed = True
        if changed:
            setattr(record, "learned_fields", learned_fields)
        return changed
