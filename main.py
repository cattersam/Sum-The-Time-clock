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
    QComboBox,
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
TEMPLATE_DIR_NAMES = ("templates", "範本")
SETTINGS_ORG = "StudentAdmin"
SETTINGS_APP = "AttendanceOCR"
INVALID_FILENAME_CHARS = r'[<>:"/\\|?*]'

COL_DAY = 0
COL_END = 1
COL_START = 2
COL_STATUS = 3
COL_SOURCE = 4

STATUS_OK = "ok"
STATUS_BLANK = "blank"
STATUS_REVIEW = "needs_review"
STATUS_FAILED = "failed"
STATUS_ABNORMAL = "abnormal_shift"
STATUS_SHORT = "short_shift"

LANGUAGES = (("en", "English"), ("zh", "中文"), ("ja", "日本語"))

TRANSLATIONS = {
    "en": {
        "app.title": "Sum The Time Clock",
        "field.name": "Name",
        "field.month": "Month",
        "field.language": "Language",
        "button.files": "Choose Photos",
        "button.template": "Choose Excel Template",
        "button.output": "Choose Output",
        "button.recognize": "Recognize",
        "button.export": "Export Excel",
        "status.ready": "Ready",
        "table.day": "Day",
        "table.end": "Clock Out",
        "table.start": "Clock In",
        "table.status": "Status",
        "table.source": "Source Image",
        "selection.none": "Not selected",
        "selection.photos": "Photos",
        "selection.template": "Excel template",
        "selection.output": "Output",
        "dialog.choose_photos": "Choose attendance sheet photos",
        "dialog.choose_template": "Choose Excel template",
        "dialog.choose_output": "Choose output file",
        "dialog.excel_filter": "Excel (*.xlsx)",
        "dialog.image_filter": "Images (*.jpg *.jpeg *.png *.bmp *.webp)",
        "dialog.error": "Error",
        "dialog.done": "Done",
        "dialog.photo_limit_title": "Photo limit",
        "dialog.photo_limit": "Only two photos can be processed at once. The first two were loaded.",
        "error.no_photos": "Choose or drop photos first.",
        "error.too_many_photos": "Only two photos can be processed at once.",
        "error.photo_pair": "Please provide one photo for days 1-15 and one for days 16-31.",
        "error.bad_drop": "Drop jpg, jpeg, png, bmp, or webp image files.",
        "error.no_template": "Choose an Excel template first.",
        "error.invalid_rows": "These days have invalid time values: {rows}",
        "status.recognizing": "Recognizing {name}",
        "status.photo_failed": "{name} recognition failed: {error}",
        "status.recognition_failed": "Recognition failed: {error}",
        "status.recognition_done": "Recognition finished. Review the table before exporting.",
        "status.learned_done": "Recognition finished. Applied {count} correction-history value(s). Review before exporting.",
        "status.exported": "Exported: {path}",
        "message.exported": "Excel exported:\n{path}",
        "status.ok": "OK",
        "status.blank": "Blank",
        "status.needs_review": "Needs Review",
        "status.failed": "Failed",
        "status.abnormal_shift": "Abnormal Shift",
        "status.short_shift": "Short Shift",
        "status.learned": "Applied from correction history",
        "tip.ok": "OK",
        "tip.blank": "No time detected.",
        "tip.failed": "OCR could not determine this row.",
        "tip.shift": "Shift duration is shorter or longer than the configured range.",
        "tip.review": "Please review this row.",
        "filename.untitled": "untitled",
        "config.error": "Configuration Error",
    },
    "zh": {
        "app.title": "打卡單辨識工具",
        "field.name": "姓名",
        "field.month": "月份",
        "field.language": "語言",
        "button.files": "選擇照片",
        "button.template": "選擇 Excel 範本",
        "button.output": "選擇輸出位置",
        "button.recognize": "開始辨識",
        "button.export": "匯出 Excel",
        "status.ready": "就緒",
        "table.day": "日期",
        "table.end": "下班",
        "table.start": "上班",
        "table.status": "狀態",
        "table.source": "來源照片",
        "selection.none": "尚未選擇",
        "selection.photos": "照片",
        "selection.template": "Excel 範本",
        "selection.output": "輸出",
        "dialog.choose_photos": "選擇打卡單照片",
        "dialog.choose_template": "選擇 Excel 範本",
        "dialog.choose_output": "選擇輸出檔案",
        "dialog.excel_filter": "Excel (*.xlsx)",
        "dialog.image_filter": "圖片 (*.jpg *.jpeg *.png *.bmp *.webp)",
        "dialog.error": "錯誤",
        "dialog.done": "完成",
        "dialog.photo_limit_title": "照片數量限制",
        "dialog.photo_limit": "一次最多處理兩張照片，已載入前兩張。",
        "error.no_photos": "請先選擇或拖入照片。",
        "error.too_many_photos": "一次最多處理兩張照片。",
        "error.photo_pair": "請提供 1-15 日與 16-31 日各一張照片。",
        "error.bad_drop": "請拖入 jpg、jpeg、png、bmp 或 webp 圖片。",
        "error.no_template": "請先選擇 Excel 範本。",
        "error.invalid_rows": "以下日期的時間格式不正確：{rows}",
        "status.recognizing": "正在辨識 {name}",
        "status.photo_failed": "{name} 辨識失敗：{error}",
        "status.recognition_failed": "辨識失敗：{error}",
        "status.recognition_done": "辨識完成，請確認表格後匯出。",
        "status.learned_done": "辨識完成，已套用 {count} 筆修正紀錄，請確認後匯出。",
        "status.exported": "已匯出：{path}",
        "message.exported": "已匯出 Excel：\n{path}",
        "status.ok": "OK",
        "status.blank": "空白",
        "status.needs_review": "需確認",
        "status.failed": "失敗",
        "status.abnormal_shift": "班段異常",
        "status.short_shift": "班段偏短",
        "status.learned": "由修正紀錄套用",
        "tip.ok": "正常。",
        "tip.blank": "沒有辨識到時間。",
        "tip.failed": "OCR 無法判斷此列。",
        "tip.shift": "工時過短或過長，請確認。",
        "tip.review": "請確認此列。",
        "filename.untitled": "未命名",
        "config.error": "設定錯誤",
    },
    "ja": {
        "app.title": "タイムカード認識ツール",
        "field.name": "氏名",
        "field.month": "月",
        "field.language": "言語",
        "button.files": "写真を選択",
        "button.template": "Excel テンプレートを選択",
        "button.output": "出力先を選択",
        "button.recognize": "認識",
        "button.export": "Excel 出力",
        "status.ready": "準備完了",
        "table.day": "日付",
        "table.end": "退勤",
        "table.start": "出勤",
        "table.status": "状態",
        "table.source": "元画像",
        "selection.none": "未選択",
        "selection.photos": "写真",
        "selection.template": "Excel テンプレート",
        "selection.output": "出力",
        "dialog.choose_photos": "タイムカード写真を選択",
        "dialog.choose_template": "Excel テンプレートを選択",
        "dialog.choose_output": "出力ファイルを選択",
        "dialog.excel_filter": "Excel (*.xlsx)",
        "dialog.image_filter": "画像 (*.jpg *.jpeg *.png *.bmp *.webp)",
        "dialog.error": "エラー",
        "dialog.done": "完了",
        "dialog.photo_limit_title": "写真数の制限",
        "dialog.photo_limit": "一度に処理できる写真は 2 枚までです。先頭の 2 枚を読み込みました。",
        "error.no_photos": "先に写真を選択またはドロップしてください。",
        "error.too_many_photos": "一度に処理できる写真は 2 枚までです。",
        "error.photo_pair": "1-15 日用と 16-31 日用の写真を 1 枚ずつ指定してください。",
        "error.bad_drop": "jpg、jpeg、png、bmp、webp の画像をドロップしてください。",
        "error.no_template": "先に Excel テンプレートを選択してください。",
        "error.invalid_rows": "次の日付の時刻形式が正しくありません: {rows}",
        "status.recognizing": "{name} を認識中",
        "status.photo_failed": "{name} の認識に失敗しました: {error}",
        "status.recognition_failed": "認識に失敗しました: {error}",
        "status.recognition_done": "認識が完了しました。出力前に表を確認してください。",
        "status.learned_done": "認識が完了しました。修正履歴から {count} 件を適用しました。",
        "status.exported": "出力済み: {path}",
        "message.exported": "Excel を出力しました:\n{path}",
        "status.ok": "OK",
        "status.blank": "空白",
        "status.needs_review": "要確認",
        "status.failed": "失敗",
        "status.abnormal_shift": "勤務時間異常",
        "status.short_shift": "短時間勤務",
        "status.learned": "修正履歴から適用",
        "tip.ok": "正常です。",
        "tip.blank": "時刻が検出されていません。",
        "tip.failed": "OCR がこの行を判定できませんでした。",
        "tip.shift": "勤務時間が設定範囲外です。",
        "tip.review": "この行を確認してください。",
        "filename.untitled": "untitled",
        "config.error": "設定エラー",
    },
}


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
        return STATUS_BLANK
    if has_raw and not has_start and not has_end:
        return STATUS_FAILED
    if has_start != has_end:
        return STATUS_REVIEW

    duration = _shift_minutes(start_time or "", end_time or "")
    if duration is None:
        return STATUS_REVIEW
    if duration < min_minutes or duration > max_minutes:
        return STATUS_ABNORMAL
    if duration < warn_short:
        return STATUS_SHORT
    if confidence and confidence < threshold:
        return STATUS_REVIEW
    return STATUS_OK


def validate_photo_pair(paths: List[Path]) -> None:
    if not paths:
        raise ImageProcessingError("no_photos")
    if len(paths) > 2:
        raise ImageProcessingError("too_many_photos")
    if len(paths) == 2:
        halves = {classify_half_by_color(read_image(path)) for path in paths}
        if halves != {"first_half", "second_half"}:
            raise ImageProcessingError("photo_pair")


def find_default_template() -> Optional[Path]:
    search_dirs = []
    for base_dir in (RUNTIME_DIR, APP_DIR):
        for dirname in TEMPLATE_DIR_NAMES:
            candidate = base_dir / dirname
            if candidate.exists():
                search_dirs.append(candidate)
    for folder in search_dirs:
        preferred = folder / "timesheet.xlsx"
        if preferred.is_file():
            return preferred
        templates = sorted(path for path in folder.glob("*.xlsx") if path.is_file())
        if templates:
            return templates[0]
    return None


def build_output_filename(month: int, name: str) -> str:
    safe_name = re.sub(INVALID_FILENAME_CHARS, "", name).strip() or "untitled"
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
                self.progress.emit(f"recognizing|{path.name}")
                try:
                    image = read_image(path)
                    fallback_half = self.config.classify_photo(path, photo_index)
                    half = classify_half_by_color(image)
                    recognized_rows = engine.recognize_attendance(image, fallback_half=half or fallback_half)
                except (ImageProcessingError, OCREngineError) as exc:
                    logging.exception("photo recognition failed")
                    fallback_half = self.config.classify_photo(path, photo_index)
                    self._mark_photo_failed(records, fallback_half, path.name)
                    self.progress.emit(f"photo_failed|{path.name}|{exc}")
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
                    records[day] = AttendanceRecord(
                        day=day,
                        start_time=start_time,
                        end_time=end_time,
                        start_raw=recognized.start_text,
                        end_raw=recognized.end_text,
                        status=status,
                        source_image=str(path),
                        confidence=recognized.confidence,
                    )

            self.finished.emit([records[day] for day in range(1, 32)])
        except (ConfigError, OCREngineError) as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            logging.exception("recognition failed")
            self.failed.emit(str(exc))

    def _mark_photo_failed(self, records: dict[int, AttendanceRecord], half: str, filename: str) -> None:
        for day in self.config.get_days(half):
            self._merge_failure(records[day], filename)

    def _merge_failure(self, current: AttendanceRecord, filename: str) -> None:
        current.status = STATUS_FAILED if current.status == STATUS_BLANK else STATUS_REVIEW
        current.source_image = "; ".join(filter(None, [current.source_image, filename]))


class MainWindow(QMainWindow):
    def __init__(self, config: AppConfig) -> None:
        super().__init__()
        self.config = config
        self.settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
        self.language = self.settings.value("language", "en", str)
        if self.language not in TRANSLATIONS:
            self.language = "en"

        self.photo_paths: List[Path] = []
        self.template_path = find_default_template()
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

        self.resize(1000, 740)
        self._build_ui()
        self.name_input.textChanged.connect(self._output_filename_source_changed)
        self.month_input.valueChanged.connect(self._output_filename_source_changed)
        self._apply_language()
        self._update_selection_label()
        self._refresh_table()

    def tr(self, key: str, **kwargs) -> str:
        text = TRANSLATIONS.get(self.language, TRANSLATIONS["en"]).get(key, TRANSLATIONS["en"].get(key, key))
        return text.format(**kwargs) if kwargs else text

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("root")
        root.setStyleSheet(
            """
            QWidget#root { background: #f5f7fb; }
            QLabel#titleLabel { font-size: 22px; font-weight: 700; color: #172033; }
            QLabel#selectionLabel { background: #ffffff; border: 2px dashed #9db3d7; border-radius: 8px; padding: 14px; color: #34415f; }
            QTableWidget { background: #ffffff; gridline-color: #e4e8f0; border: 1px solid #dce3ef; }
            QPushButton { padding: 7px 12px; border-radius: 5px; background: #2457a6; color: white; }
            QPushButton:disabled { background: #a9b6ca; }
            QLineEdit, QSpinBox, QComboBox { padding: 5px; border: 1px solid #cbd5e1; border-radius: 4px; background: white; }
            """
        )
        layout = QVBoxLayout(root)
        layout.setSpacing(10)

        self.title_label = QLabel()
        self.title_label.setObjectName("titleLabel")
        layout.addWidget(self.title_label)

        form = QFormLayout()
        self.name_label = QLabel()
        self.month_label = QLabel()
        self.language_label = QLabel()
        self.name_input = QLineEdit()
        self.month_input = QSpinBox()
        self.month_input.setRange(1, 12)
        self.month_input.setValue(get_default_month())
        self.language_combo = QComboBox()
        for code, label in LANGUAGES:
            self.language_combo.addItem(label, code)
        index = self.language_combo.findData(self.language)
        self.language_combo.setCurrentIndex(max(index, 0))
        self.language_combo.currentIndexChanged.connect(self._language_changed)
        form.addRow(self.name_label, self.name_input)
        form.addRow(self.month_label, self.month_input)
        form.addRow(self.language_label, self.language_combo)
        layout.addLayout(form)

        button_row = QHBoxLayout()
        self.files_button = QPushButton()
        self.template_button = QPushButton()
        self.output_button = QPushButton()
        for button in (self.files_button, self.template_button, self.output_button):
            button_row.addWidget(button)
        layout.addLayout(button_row)

        if self.config.ui.get("enable_drag_drop", True):
            self.selection_label = DropImageLabel()
            self.selection_label.filesDropped.connect(self._set_photo_paths_from_drop)
        else:
            self.selection_label = QLabel()
        self.selection_label.setObjectName("selectionLabel")
        self.selection_label.setWordWrap(True)
        layout.addWidget(self.selection_label)

        action_row = QHBoxLayout()
        self.recognize_button = QPushButton()
        self.export_button = QPushButton()
        action_row.addWidget(self.recognize_button)
        action_row.addWidget(self.export_button)
        layout.addLayout(action_row)

        self.status_label = QLabel()
        layout.addWidget(self.status_label)

        self.table = QTableWidget(31, 5)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table)

        self.setCentralWidget(root)
        self.files_button.clicked.connect(self._select_files)
        self.template_button.clicked.connect(self._select_template)
        self.output_button.clicked.connect(self._select_output)
        self.recognize_button.clicked.connect(self._start_recognition)
        self.export_button.clicked.connect(self._export)
        self.table.cellChanged.connect(self._cell_changed)

    def _apply_language(self) -> None:
        self.setWindowTitle(self.tr("app.title"))
        self.title_label.setText(self.tr("app.title"))
        self.name_label.setText(self.tr("field.name"))
        self.month_label.setText(self.tr("field.month"))
        self.language_label.setText(self.tr("field.language"))
        self.files_button.setText(self.tr("button.files"))
        self.template_button.setText(self.tr("button.template"))
        self.output_button.setText(self.tr("button.output"))
        self.recognize_button.setText(self.tr("button.recognize"))
        self.export_button.setText(self.tr("button.export"))
        self.table.setHorizontalHeaderLabels(
            (
                self.tr("table.day"),
                self.tr("table.end"),
                self.tr("table.start"),
                self.tr("table.status"),
                self.tr("table.source"),
            )
        )
        if not self.status_label.text():
            self.status_label.setText(self.tr("status.ready"))

    def _language_changed(self, *_args: object) -> None:
        code = self.language_combo.currentData()
        if code not in TRANSLATIONS:
            return
        self.language = code
        self.settings.setValue("language", self.language)
        self._apply_language()
        self._update_selection_label()
        self._refresh_table()

    def _translate_error(self, exc: Exception) -> str:
        mapping = {
            "no_photos": "error.no_photos",
            "too_many_photos": "error.too_many_photos",
            "photo_pair": "error.photo_pair",
        }
        message = str(exc)
        return self.tr(mapping.get(message, message))

    def _select_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            self.tr("dialog.choose_photos"),
            "",
            self.tr("dialog.image_filter"),
        )
        self._set_photo_paths(filter_image_files(paths))

    def _set_photo_paths_from_drop(self, paths: list[Path]) -> None:
        image_paths = filter_image_files(paths)
        if not image_paths:
            self._show_error(self.tr("error.bad_drop"))
            return
        if len(image_paths) > 2:
            QMessageBox.information(self, self.tr("dialog.photo_limit_title"), self.tr("dialog.photo_limit"))
            image_paths = image_paths[:2]
        self._set_photo_paths(image_paths)

    def _set_photo_paths(self, paths: List[Path]) -> None:
        try:
            validate_photo_pair(paths)
        except ImageProcessingError as exc:
            self._show_error(self._translate_error(exc))
            return
        self.photo_paths = list(paths)
        self._update_selection_label()

    def _select_template(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, self.tr("dialog.choose_template"), "", self.tr("dialog.excel_filter"))
        if path:
            self.template_path = Path(path)
            self._update_selection_label()

    def _select_output(self) -> None:
        default_path = self._suggested_output_path()
        path, _ = QFileDialog.getSaveFileName(self, self.tr("dialog.choose_output"), str(default_path), self.tr("dialog.excel_filter"))
        if path:
            self.output_path = Path(path).with_suffix(".xlsx")
            self.output_dir = self.output_path.parent
            self._manual_output_path = True
            self.settings.setValue("last_output_dir", str(self.output_dir))
            self._update_selection_label()

    def _update_selection_label(self) -> None:
        template = self.template_path.name if self.template_path else self.tr("selection.none")
        output_path = self._current_output_path()
        photos = ", ".join(path.name for path in self.photo_paths) if self.photo_paths else self.tr("selection.none")
        self.selection_label.setText(
            f"{self.tr('selection.photos')}: {photos}\n"
            f"{self.tr('selection.template')}: {template}\n"
            f"{self.tr('selection.output')}: {output_path}"
        )

    def _load_output_dir(self) -> Path:
        saved = self.settings.value("last_output_dir", "", str)
        path = Path(saved) if saved else RUNTIME_DIR / "output"
        return path if path.exists() else RUNTIME_DIR / "output"

    def _suggested_output_path(self) -> Path:
        filename = build_output_filename(self.month_input.value(), self.name_input.text())
        if filename.startswith("untitled") or filename.endswith("_untitled.xlsx"):
            filename = filename.replace("untitled", self.tr("filename.untitled"))
        return self.output_dir / filename

    def _current_output_path(self) -> Path:
        if self._manual_output_path and self.output_path:
            return self.output_path
        return self._suggested_output_path()

    def _output_filename_source_changed(self, *_args: object) -> None:
        self._manual_output_path = False
        self._update_selection_label()

    def _start_recognition(self) -> None:
        if not self.photo_paths:
            self._show_error(self.tr("error.no_photos"))
            return
        try:
            validate_photo_pair(self.photo_paths)
        except ImageProcessingError as exc:
            self._show_error(self._translate_error(exc))
            return
        self._set_busy(True)
        self.thread = QThread()
        self.worker = RecognitionWorker(self.photo_paths, self.config)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self._recognition_finished)
        self.worker.failed.connect(self._recognition_failed)
        self.worker.progress.connect(self._recognition_progress)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self._thread_finished)
        self.thread.start()

    def _recognition_progress(self, payload: str) -> None:
        parts = payload.split("|", 2)
        if parts[0] == "recognizing" and len(parts) > 1:
            self.status_label.setText(self.tr("status.recognizing", name=parts[1]))
        elif parts[0] == "photo_failed" and len(parts) > 2:
            self.status_label.setText(self.tr("status.photo_failed", name=parts[1], error=parts[2]))
        else:
            self.status_label.setText(payload)

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
            self.status_label.setText(self.tr("status.learned_done", count=learned_count))
        else:
            self.status_label.setText(self.tr("status.recognition_done"))

    def _recognition_failed(self, message: str) -> None:
        self.status_label.setText(self.tr("status.recognition_failed", error=message))
        self._show_error(self.tr("status.recognition_failed", error=message))

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
        status = self.tr(f"status.{record.status}")
        if record.learned_fields:
            return f"{status}; {self.tr('status.learned')}"
        return status

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
        if status == STATUS_OK:
            bg = QColor(224, 242, 254) if learned else QColor(255, 255, 255)
            tip = self.tr("status.learned") if learned else self.tr("tip.ok")
        elif status == STATUS_BLANK:
            bg = QColor(245, 245, 245)
            tip = self.tr("tip.blank")
        elif status == STATUS_FAILED:
            bg = QColor(255, 210, 210)
            tip = self.tr("tip.failed")
        elif status in {STATUS_ABNORMAL, STATUS_SHORT}:
            bg = QColor(255, 225, 190)
            tip = self.tr("tip.shift")
        else:
            bg = QColor(255, 243, 176)
            tip = self.tr("tip.review")
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
            self._show_error(self.tr("error.no_template"))
            return
        invalid_rows = []
        for row in range(self.table.rowCount()):
            for column in (COL_END, COL_START):
                text = self.table.item(row, column).text().strip() if self.table.item(row, column) else ""
                if text and not extract_time_text(text):
                    invalid_rows.append(str(row + 1))
        if invalid_rows:
            self._show_error(self.tr("error.invalid_rows", rows=", ".join(sorted(set(invalid_rows), key=int))))
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
            self.status_label.setText(self.tr("status.exported", path=output_path))
            QMessageBox.information(self, self.tr("dialog.done"), self.tr("message.exported", path=output_path))
        except ExcelExportError as exc:
            self._show_error(str(exc))

    def _show_error(self, message: str) -> None:
        logging.error(message)
        QMessageBox.critical(self, self.tr("dialog.error"), message)


def main() -> int:
    app = QApplication(sys.argv)
    setup_logging(RUNTIME_DIR / "logs")
    try:
        config = AppConfig(find_config_path())
    except ConfigError as exc:
        logging.exception("config failed")
        QMessageBox.critical(None, TRANSLATIONS["en"]["config.error"], str(exc))
        return 1
    window = MainWindow(config)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
