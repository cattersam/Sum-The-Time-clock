from __future__ import annotations

from pathlib import Path
import tempfile

from correction_history import CorrectionHistory
from excel_writer import export_excel
from models import AttendanceRecord
from roi_config import AppConfig
from utils import extract_time_text, filter_image_files
from main import validate_photo_pair


def main() -> None:
    assert extract_time_text("O8.3O") == "08:30"
    assert extract_time_text("18;05") == "18:05"
    assert extract_time_text("18：05") == "18:05"

    config = AppConfig("config.json")
    assert config.validation["min_shift_minutes"] == 120
    assert config.learning["min_count"] == 2

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        image = tmp_path / "a.webp"
        first = tmp_path / "first.jpg"
        second = tmp_path / "second.jpg"
        image.write_bytes(b"fake")
        first.write_bytes(b"fake")
        second.write_bytes(b"fake")
        assert filter_image_files([image]) == [image]
        validate_photo_pair([first, second])

        history = CorrectionHistory(tmp_path / "correction_history.jsonl", min_count=2)
        for _ in range(2):
            assert history.append(
                date=1,
                field="clock_in",
                ocr_text="O8.3O",
                old_value="",
                new_value="08:30",
                image_path="sample.jpg",
                confidence=0.5,
            )
        record = AttendanceRecord(day=1, start_raw="O8.3O", end_raw="17:30", end_time="17:30")
        assert history.apply_to_record(record)
        assert record.start_time == "08:30"
        assert "clock_in" in record.learned_fields

        output = tmp_path / "out.xlsx"
        records = [AttendanceRecord(day=1, start_time="08:30", end_time="17:30")]
        export_excel("templates/timesheet.xlsx", output, "Test", 7, records, config.excel_mapping)
        assert output.exists() and output.stat().st_size > 0

    print("static logic tests passed")


if __name__ == "__main__":
    main()
