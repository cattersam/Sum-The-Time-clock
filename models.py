from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional, Set


@dataclass
class OCRResult:
    text: str = ""
    confidence: float = 0.0


@dataclass
class OCRAttendanceRow:
    day: int
    start_text: str = ""
    end_text: str = ""
    confidence: float = 0.0


@dataclass
class AttendanceRecord:
    day: int
    start_time: Optional[str] = None
    end_time: Optional[str] = None
    status: str = "blank"
    source_image: str = ""
    confidence: float = 0.0
    start_raw: str = ""
    end_raw: str = ""
    learned_fields: Set[str] = field(default_factory=set)
