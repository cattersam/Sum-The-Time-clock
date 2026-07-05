from __future__ import annotations

from pathlib import Path
import os
import tempfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import cv2
import numpy as np
from PySide6.QtWidgets import QApplication, QMessageBox

import main


def write_color(path: Path, bgr: tuple[int, int, int]) -> None:
    image = np.zeros((40, 40, 3), dtype=np.uint8)
    image[:, :] = bgr
    ok, encoded = cv2.imencode(path.suffix, image)
    assert ok
    path.write_bytes(encoded.tobytes())


def main_test() -> None:
    app = QApplication.instance() or QApplication([])
    config = main.AppConfig("config.json")
    window = main.MainWindow(config)
    QMessageBox.information = lambda *args, **kwargs: QMessageBox.StandardButton.Ok
    QMessageBox.critical = lambda *args, **kwargs: QMessageBox.StandardButton.Ok

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        first = tmp_path / "first.jpg"
        second = tmp_path / "second.png"
        third = tmp_path / "third.webp"
        write_color(first, (0, 0, 255))
        write_color(second, (0, 255, 0))
        write_color(third, (0, 0, 255))

        window._set_photo_paths_from_drop([first])
        assert [path.name for path in window.photo_paths] == ["first.jpg"]

        window._set_photo_paths_from_drop([first, second])
        assert [path.name for path in window.photo_paths] == ["first.jpg", "second.png"]

        window._set_photo_paths_from_drop([first, second, third])
        assert [path.name for path in window.photo_paths] == ["first.jpg", "second.png"]

    record = window.records[0]
    record.start_raw = "O8.3O"
    record.start_time = None
    record.end_time = "17:30"
    record.end_raw = "17:30"
    record.source_image = "sample.jpg"
    record.confidence = 0.7
    window._refresh_table()
    before = window.correction_history.path.stat().st_size if window.correction_history.path.exists() else 0
    item = window.table.item(0, main.COL_START)
    item.setText("08:30")
    window._cell_changed(0, main.COL_START)
    after = window.correction_history.path.stat().st_size
    assert after > before
    assert window.records[0].start_time == "08:30"
    assert window.table.item(0, main.COL_STATUS).text()

    app.quit()
    print("gui logic tests passed")


if __name__ == "__main__":
    main_test()
