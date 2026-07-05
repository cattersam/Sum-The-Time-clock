from __future__ import annotations

from pathlib import Path
from typing import List, Optional
import logging
import re
import sys

from PySide6.QtCore import QObject, QSettings, QThread, Qt, Signal
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from correction_history import CorrectionHistory
from drag_drop_widgets import DropImageLabel
from excel_writer import ExcelExportError, export_excel
from image_preprocess import ImageProcessingError, classify_half_by_color, read_image
from models import AttendanceRecord
from ocr_engine import OCREngine, OCREngineError
from roi_config import AppConfig, ConfigError
from utils import extract_time_text, filter_image_files, get_default_month, setup_logging


APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
RUNTIME_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
TEMPLATE_DIR_NAMES = ("範本", "templates")
SETTINGS_ORG = "StudentAdmin"
SETTINGS_APP = "AttendanceOCR"
INVALID_FILENAME_CHARS = r'[<>:"/\\|?*]'

COL_DAY = 0
COL_END = 1
COL_START = 2
COL_STATUS = 3
COL_SOURCE = 4


def _parse_minutes(value: str) -> Optional[int]:
    normalized = extract_time_text(value or "")
    if not normalized:
        return None
    hours, minutes = (int(part) for part in normalized.split(":"))
    return hours * 60 + minutes


def _shift_minutes(start_time: str, end_time: str) -> Optional[int]:
    start = _parse_minutes(start_time)
    end = _parse_minutes(end_time)
    if start is None or end is None:
        return None
    duration = end - start
    if duration < 0:
        return None
    return duration


def record_status(
    start_raw: str,
    end_raw: str,
    start_time: Optional[str],
    end_time: Optional[str],
    confidence: float,
    validation: Optional[dict] = None,
) -> str:
    validation = validation or {}
    min_minutes = int(validation.get("min_shift_minutes", 120))
    max_minutes = int(validation.get("max_shift_minutes", 960))
    warn_short = int(validation.get("warn_short_shift_minutes", 240))
    threshold = float(validation.get("ok_confidence_threshold", 0.88))

    has_raw = bool(str(start_raw or "").strip() or str(end_raw or "").strip())
    has_start = bool(start_time)
    has_end = bool(end_time)
    if not has_raw and not has_start and not has_end:
        return "空白"
    if has_raw and not has_start and not has_end:
        return "失敗"
    if has_start != has_end:
        return "需確認"

    duration = _shift_minutes(start_time or "", end_time or "")
    if duration is None:
        return "需確認"
    if duration < min_minutes or duration > max_minutes:
        return "班段異常"
    if duration < warn_short:
        return "班段偏短"
    if confidence and confidence < threshold:
        return "需確認"
    return "OK"


def validate_photo_pair(paths: List[Path]) -> None:
    if not paths:
        raise ImageProcessingError("請先選擇或拖入照片")
    if len(paths) > 2:
        raise ImageProcessingError("一次最多處理 2 張照片")
    if len(paths) == 2:
        halves = {classify_half_by_color(read_image(path)) for path in paths}
        if halves != {"first_half", "second_half"}:
            raise ImageProcessingError("請確認兩張照片分別是 1-15 日與 16-31 日")


def find_default_template() -> Optional[Path]:
    search_dirs = []
    for base_dir in (RUNTIME_DIR, APP_DIR):
        for dirname in TEMPLATE_DIR_NAMES:
            candidate = base_dir / dirname
            if candidate.exists():
                search_dirs.append(candidate)
    for folder in search_dirs:
        templates = sorted(path for path in folder.glob("*.xlsx") if path.is_file())
        if templates:
            return templates[0]
    return None


def build_output_filename(month: int, name: str) -> str:
    safe_name = re.sub(INVALID_FILENAME_CHARS, "", name).strip() or "未命名"
    return f"{month:02d}_{safe_name}.xlsx"


def find_config_path() -> Path:
    runtime_config = RUNTIME_DIR / "config.json"
    if runtime_config.exists():
        return runtime_config
    return APP_DIR / "config.json"


class RecognitionWorker(QObject):
    finished = Signal(list)
    failed = Signal(str)
    progress = Signal(str)

    def __init__(self, paths: List[Path], config: AppConfig) -> None:
        super().__init__()
        self.paths = paths
        self.config = config

    def run(self) -> None:
        try:
            records = {day: AttendanceRecord(day=day) for day in range(1, 32)}
            engine = OCREngine()
            engine.initialize()
            for photo_index, path in enumerate(self.paths):
                self.progress.emit(f"辨識 {path.name}")
                try:
                    image = read_image(path)
                    fallback_half = self.config.classify_photo(path, photo_index)
                    half = classify_half_by_color(image)
                    recognized_rows = engine.recognize_attendance(image, fallback_half=half or fallback_half)
                except (ImageProcessingError, OCREngineError) as exc:
                    logging.exception("photo recognition failed")
                    fallback_half = self.config.classify_photo(path, photo_index)
                    self._mark_photo_failed(records, fallback_half, path.name)
                    self.progress.emit(f"{path.name} 辨識失敗: {exc}")
                    continue

                for recognized in recognized_rows:
                    day = recognized.day
                    start_time = extract_time_text(recognized.start_text or "")
                    end_time = extract_time_text(recognized.end_text or "")
                    status = record_status(
                        recognized.start_text,
                        recognized.end_text,
                        start_time,
                        end_time,
                        recognized.confidence,
                        self.config.validation,
                    )
                    incoming = AttendanceRecord(
                        day=day,
                        start_time=start_time,
                        end_time=end_time,
                        start_raw=recognized.start_text,
                        end_raw=recognized.end_text,
                        status=status,
                        source_image=str(path),
                        confidence=recognized.confidence,
                    )
                    current = records.get(day)
                    if current and current.status != "空白":
                        current.source_image = "; ".join(filter(None, [current.source_image, str(path)]))
                    records[day] = incoming

            self.finished.emit([records[day] for day in range(1, 32)])
        except (ConfigError, OCREngineError) as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            logging.exception("recognition failed")
            self.failed.emit(f"辨識失敗: {exc}")

    def _mark_photo_failed(self, records: dict[int, AttendanceRecord], half: str, filename: str) -> None:
        for day in self.config.get_days(half):
            self._merge_failure(records[day], filename)

    def _merge_failure(self, current: AttendanceRecord, filename: str) -> None:
        current.status = "失敗" if current.status == "空白" else "需確認"
        current.source_image = "; ".join(filter(None, [current.source_image, filename]))


class MainWindow(QMainWindow):
    def __init__(self, config: AppConfig) -> None:
        super().__init__()
        self.config = config
        self.photo_paths: List[Path] = []
        self.template_path = find_default_template()
        self.settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
        self.output_dir = self._load_output_dir()
        self.output_path: Optional[Path] = None
        self._manual_output_path = False
        self.records = [AttendanceRecord(day=day) for day in range(1, 32)]
        self.thread: Optional[QThread] = None
        self.worker: Optional[RecognitionWorker] = None
        self._updating_table = False
        self._last_table_values: dict[tuple[int, int], str] = {}

        learning = self.config.learning
        self.correction_history = CorrectionHistory(
            RUNTIME_DIR / str(learning.get("history_file", "logs/correction_history.jsonl")),
            min_count=int(learning.get("min_count", 2)),
        )

        self.setWindowTitle("打卡單辨識工具")
        self.resize(920, 720)
        self._build_ui()
        self.name_input.textChanged.connect(self._output_filename_source_changed)
        self.month_input.valueChanged.connect(self._output_filename_source_changed)
        self._update_selection_label()
        self._refresh_table()

    def _build_ui(self) -> None:
        root = QWidget()
        layout = QVBoxLayout(root)

        form = QFormLayout()
        self.name_input = QLineEdit()
        self.month_input = QSpinBox()
        self.month_input.setRange(1, 12)
        self.month_input.setValue(get_default_month())
        form.addRow("姓名", self.name_input)
        form.addRow("月份", self.month_input)
        layout.addLayout(form)

        button_row = QHBoxLayout()
        self.files_button = QPushButton("選擇照片")
        self.template_button = QPushButton("選擇 Excel 範本")
        self.output_button = QPushButton("選擇輸出位置")
        for button in (self.files_button, self.template_button, self.output_button):
            button_row.addWidget(button)
        layout.addLayout(button_row)

        if self.config.ui.get("enable_drag_drop", True):
            self.selection_label = DropImageLabel()
            self.selection_label.filesDropped.connect(self._set_photo_paths_from_drop)
        else:
            self.selection_label = QLabel()
        self.selection_label.setWordWrap(True)
        layout.addWidget(self.selection_label)

        action_row = QHBoxLayout()
        self.recognize_button = QPushButton("開始辨識")
        self.export_button = QPushButton("匯出 Excel")
        action_row.addWidget(self.recognize_button)
        action_row.addWidget(self.export_button)
        layout.addLayout(action_row)

        self.status_label = QLabel("就緒")
        layout.addWidget(self.status_label)

        self.table = QTableWidget(31, 5)
        self.table.setHorizontalHeaderLabels(("日期", "下班", "上班", "狀態", "來源照片"))
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table)

        self.setCentralWidget(root)
        self.files_button.clicked.connect(self._select_files)
        self.template_button.clicked.connect(self._select_template)
        self.output_button.clicked.connect(self._select_output)
        self.recognize_button.clicked.connect(self._start_recognition)
        self.export_button.clicked.connect(self._export)
        self.table.cellChanged.connect(self._cell_changed)

    def _select_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "選擇打卡單照片",
            "",
            "圖片 (*.jpg *.jpeg *.png *.bmp *.webp)",
        )
        self._set_photo_paths(filter_image_files(paths))

    def _set_photo_paths_from_drop(self, paths: list[Path]) -> None:
        image_paths = filter_image_files(paths)
        if not image_paths:
            self._show_error("請拖入 jpg、jpeg、png、bmp 或 webp 圖片")
            return
        if len(image_paths) > 2:
            QMessageBox.information(self, "照片數量限制", "一次最多處理兩張照片，已先取前兩張。")
            image_paths = image_paths[:2]
        self._set_photo_paths(image_paths)

    def _set_photo_paths(self, paths: List[Path]) -> None:
        try:
            validate_photo_pair(paths)
        except ImageProcessingError as exc:
            self._show_error(str(exc))
            return
        self.photo_paths = list(paths)
        self._update_selection_label()

    def _select_template(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "選擇 Excel 範本", "", "Excel (*.xlsx)")
        if path:
            self.template_path = Path(path)
            self._update_selection_label()

    def _select_output(self) -> None:
        default_path = self._suggested_output_path()
        path, _ = QFileDialog.getSaveFileName(self, "選擇輸出位置", str(default_path), "Excel (*.xlsx)")
        if path:
            self.output_path = Path(path).with_suffix(".xlsx")
            self.output_dir = self.output_path.parent
            self._manual_output_path = True
            self.settings.setValue("last_output_dir", str(self.output_dir))
            self._update_selection_label()

    def _update_selection_label(self) -> None:
        template = self.template_path.name if self.template_path else "尚未選擇"
        output_path = self._current_output_path()
        photos = "、".join(path.name for path in self.photo_paths) if self.photo_paths else "尚未選擇"
        self.selection_label.setText(
            f"照片：{photos}\nExcel 範本：{template}\n輸出：{output_path}"
        )

    def _load_output_dir(self) -> Path:
        saved = self.settings.value("last_output_dir", "", str)
        path = Path(saved) if saved else RUNTIME_DIR / "output"
        return path if path.exists() else RUNTIME_DIR / "output"

    def _suggested_output_path(self) -> Path:
        return self.output_dir / build_output_filename(self.month_input.value(), self.name_input.text())

    def _current_output_path(self) -> Path:
        if self._manual_output_path and self.output_path:
            return self.output_path
        return self._suggested_output_path()

    def _output_filename_source_changed(self, *_args: object) -> None:
        self._manual_output_path = False
        self._update_selection_label()

    def _start_recognition(self) -> None:
        if not self.photo_paths:
            self._show_error("請先選擇或拖入照片")
            return
        try:
            validate_photo_pair(self.photo_paths)
        except ImageProcessingError as exc:
            self._show_error(str(exc))
            return
        self._set_busy(True)
        self.thread = QThread()
        self.worker = RecognitionWorker(self.photo_paths, self.config)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self._recognition_finished)
        self.worker.failed.connect(self._recognition_failed)
        self.worker.progress.connect(self.status_label.setText)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self._thread_finished)
        self.thread.start()

    def _recognition_finished(self, records: List[AttendanceRecord]) -> None:
        learned_count = 0
        if self.config.learning.get("enabled", True):
            for record in records:
                if self.correction_history.apply_to_record(record):
                    learned_count += 1
                record.status = record_status(
                    record.start_raw,
                    record.end_raw,
                    record.start_time,
                    record.end_time,
                    record.confidence,
                    self.config.validation,
                )
        self.records = records
        self._refresh_table()
        if learned_count:
            self.status_label.setText(f"辨識完成，已由修正紀錄套用 {learned_count} 筆，請確認後匯出 Excel")
        else:
            self.status_label.setText("辨識完成，請確認後匯出 Excel")

    def _recognition_failed(self, message: str) -> None:
        self.status_label.setText(message)
        self._show_error(message)

    def _thread_finished(self) -> None:
        self._set_busy(False)
        if self.worker:
            self.worker.deleteLater()
        self.worker = None
        self.thread = None

    def _set_busy(self, busy: bool) -> None:
        for button in (self.files_button, self.template_button, self.output_button, self.recognize_button, self.export_button):
            button.setEnabled(not busy)

    def _display_status(self, record: AttendanceRecord) -> str:
        if record.learned_fields:
            return f"{record.status}；由修正紀錄套用"
        return record.status

    def _refresh_table(self) -> None:
        self._updating_table = True
        try:
            for row, record in enumerate(self.records):
                values = [
                    str(record.day),
                    record.end_time or "",
                    record.start_time or "",
                    self._display_status(record),
                    Path(record.source_image).name if record.source_image else "",
                ]
                for column, value in enumerate(values):
                    item = QTableWidgetItem(value)
                    if column in (COL_DAY, COL_STATUS, COL_SOURCE):
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self.table.setItem(row, column, item)
                    self._last_table_values[(row, column)] = value
                self._apply_row_style(row, record)
        finally:
            self._updating_table = False

    def _apply_row_style(self, row: int, record: AttendanceRecord) -> None:
        if not self.config.ui.get("highlight_abnormal_rows", True):
            return
        status = record.status or ""
        learned = bool(record.learned_fields)
        if status == "OK":
            bg = QColor(224, 242, 254) if learned else QColor(255, 255, 255)
            tip = "由修正紀錄套用" if learned else "OK"
        elif status == "空白":
            bg = QColor(245, 245, 245)
            tip = "空白"
        elif status == "失敗":
            bg = QColor(255, 210, 210)
            tip = "OCR 無法判斷"
        elif status in {"班段異常", "班段偏短"}:
            bg = QColor(255, 225, 190)
            tip = "工時過短或過長，請確認"
        else:
            bg = QColor(255, 243, 176)
            tip = "需確認"
        for column in range(self.table.columnCount()):
            item = self.table.item(row, column)
            if item is not None:
                item.setBackground(QBrush(bg))
                item.setToolTip(tip)

    def _cell_changed(self, row: int, column: int) -> None:
        if self._updating_table or row >= len(self.records) or column not in (COL_END, COL_START):
            return
        item = self.table.item(row, column)
        new_value = extract_time_text(item.text().strip()) if item else None
        record = self.records[row]
        old_value = self._last_table_values.get((row, column), "")
        typed_value = item.text().strip() if item else ""
        normalized_value = new_value or typed_value
        if column == COL_START:
            raw_text = record.start_raw or old_value
            record.start_time = new_value
            field = "clock_in"
        else:
            raw_text = record.end_raw or old_value
            record.end_time = new_value
            field = "clock_out"

        if old_value != normalized_value and normalized_value:
            self.correction_history.append(
                date=record.day,
                field=field,
                ocr_text=raw_text,
                old_value=old_value,
                new_value=normalized_value,
                image_path=record.source_image,
                confidence=record.confidence,
            )
            self._last_table_values[(row, column)] = normalized_value

        record.status = record_status(
            record.start_raw,
            record.end_raw,
            record.start_time,
            record.end_time,
            record.confidence,
            self.config.validation,
        )
        if item and new_value and item.text() != new_value:
            self._updating_table = True
            item.setText(new_value)
            self._updating_table = False
        status_item = self.table.item(row, COL_STATUS)
        if status_item:
            status_item.setText(self._display_status(record))
        self._apply_row_style(row, record)

    def _export(self) -> None:
        if not self.template_path:
            self._show_error("請先選擇 Excel 範本")
            return
        invalid_rows = []
        for row in range(self.table.rowCount()):
            for column in (COL_END, COL_START):
                text = self.table.item(row, column).text().strip() if self.table.item(row, column) else ""
                if text and not extract_time_text(text):
                    invalid_rows.append(str(row + 1))
        if invalid_rows:
            self._show_error(f"以下日期時間格式不正確: {', '.join(sorted(set(invalid_rows), key=int))}")
            return

        output_path = self._current_output_path()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            export_excel(
                self.template_path,
                output_path,
                self.name_input.text(),
                self.month_input.value(),
                self.records,
                self.config.excel_mapping,
            )
            self.output_dir = output_path.parent
            self.settings.setValue("last_output_dir", str(self.output_dir))
            self.status_label.setText(f"已匯出: {output_path}")
            QMessageBox.information(self, "完成", f"已匯出 Excel:\n{output_path}")
        except ExcelExportError as exc:
            self._show_error(str(exc))

    def _show_error(self, message: str) -> None:
        logging.error(message)
        QMessageBox.critical(self, "錯誤", message)


def main() -> int:
    app = QApplication(sys.argv)
    setup_logging(RUNTIME_DIR / "logs")
    try:
        config = AppConfig(find_config_path())
    except ConfigError as exc:
        logging.exception("config failed")
        QMessageBox.critical(None, "設定錯誤", str(exc))
        return 1
    window = MainWindow(config)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
