"""Small PySide6 drag-and-drop widgets for Attendance OCR."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QLabel

IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"}


class DropImageLabel(QLabel):
    """A QLabel that accepts dropped image files and emits Path objects."""

    filesDropped = Signal(list)

    def __init__(self, text: str = "") -> None:
        super().__init__(text)
        self.setAcceptDrops(True)
        self.setStyleSheet(
            "QLabel { border: 2px dashed #999; border-radius: 8px; padding: 10px; }"
            "QLabel:hover { border-color: #555; }"
        )

    def dragEnterEvent(self, event):  # noqa: N802 - Qt API name
        urls = event.mimeData().urls() if event.mimeData().hasUrls() else []
        if any(Path(url.toLocalFile()).suffix.lower() in IMAGE_SUFFIXES for url in urls):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):  # noqa: N802 - Qt API name
        paths = []
        for url in event.mimeData().urls():
            p = Path(url.toLocalFile())
            if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES:
                paths.append(p)
        if paths:
            self.filesDropped.emit(paths)
            event.acceptProposedAction()
        else:
            event.ignore()
