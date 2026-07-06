from __future__ import annotations

from datetime import datetime
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
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from correction_history import CorrectionHistory
from drag_drop_widgets import DropImageLabel
from excel_writer import ExcelExportError, export_excel
from image_preprocess import ImageProcessingError, read_image
from models import AttendanceRecord
from ocr_engine import OCREngine, OCREngineError
from roi_config import AppConfig, ConfigError
from utils import extract_time_text, filter_image_files, get_default_month, setup_logging


APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
RUNTIME_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
TEMPLATE_DIR_NAMES = ("templates",)
SETTINGS_ORG = "StudentAdmin"
SETTINGS_APP = "SumTheTimeClock"
INVALID_FILENAME_CHARS = r'[<>:"/\\|?*]'

COL_INDEX = 0
COL_DATE = 1
COL_START = 2
COL_END = 3
COL_TOTAL = 4
COL_STATUS = 5
COL_SOURCE = 6

STATUS_OK = "ok"
STATUS_BLANK = "blank"
STATUS_REVIEW = "needs_review"
STATUS_FAILED = "failed"
STATUS_ABNORMAL = "abnormal_shift"
STATUS_SHORT = "short_shift"

LANGUAGES = (("en", "English"), ("zh", "繁體中文"), ("ja", "日本語"))

TRANSLATIONS = {
    "en": {
        "app.title": "Sum The Time Clock",
        "project.section": "PROJECT",
        "workflow.section": "WORKFLOW",
        "info.section": "PROJECT INFO",
        "field.name": "Name",
        "field.month": "Month",
        "field.language": "Language",
        "placeholder.name": "Employee or project name",
        "button.new_project": "New Project",
        "button.files": "Choose Photos",
        "button.template": "Template",
        "button.output": "Output",
        "button.recognize": "Recognize",
        "button.export": "Export Excel",
        "button.settings": "Project Settings",
        "drop.title": "Drop images here",
        "drop.subtitle": "Drag and drop time clock images to get started",
        "drop.support": "Supports JPG, JPEG, PNG, BMP, WEBP",
        "drop.loaded": "Loaded images",
        "status.ready": "Ready",
        "status.records": "Records: {count}",
        "status.images": "Images: {count}",
        "status.language": "Language: {language}",
        "table.index": "#",
        "table.date": "Date",
        "table.start": "Time In",
        "table.end": "Time Out",
        "table.total": "Total Hours",
        "table.status": "Status",
        "table.source": "Source",
        "selection.none": "Not selected",
        "selection.template": "Excel template",
        "selection.output": "Output",
        "info.created": "Created:",
        "info.images": "Images:",
        "info.records": "Records:",
        "info.updated": "Last Updated:",
        "workflow.import.title": "Import Images",
        "workflow.import.body": "Drag and drop time clock images above",
        "workflow.recognize.title": "Recognize",
        "workflow.recognize.body": "Extract attendance data using OCR",
        "workflow.review.title": "Review & Edit",
        "workflow.review.body": "Verify and correct recognized data",
        "workflow.export.title": "Export",
        "workflow.export.body": "Export attendance to Excel",
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
        "progress.loading_engine": "Loading OCR engine",
        "progress.engine_ready": "OCR engine ready",
        "progress.recognizing": "Recognizing {name}",
        "progress.finished_photo": "Finished {name}",
        "progress.done": "Recognition finished",
        "status.photo_failed": "{name} recognition failed: {error}",
        "status.recognition_failed": "Recognition failed: {error}",
        "status.recognition_done": "Recognition finished. Review the table before exporting.",
        "status.learned_done": "Recognition finished. Applied {count} correction-history value(s). Review before exporting.",
        "status.exported": "Exported: {path}",
        "message.exported": "Excel exported:\n{path}",
        "status.ok": "Present",
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
        "app.title": "時數加總打卡工具",
        "project.section": "專案",
        "workflow.section": "流程",
        "info.section": "專案資訊",
        "field.name": "姓名",
        "field.month": "月份",
        "field.language": "語言",
        "placeholder.name": "員工或專案名稱",
        "button.new_project": "新專案",
        "button.files": "選擇照片",
        "button.template": "範本",
        "button.output": "輸出",
        "button.recognize": "開始辨識",
        "button.export": "匯出 Excel",
        "button.settings": "專案設定",
        "drop.title": "把圖片拖到這裡",
        "drop.subtitle": "拖放打卡單照片即可開始",
        "drop.support": "支援 JPG、JPEG、PNG、BMP、WEBP",
        "drop.loaded": "已載入圖片",
        "status.ready": "就緒",
        "status.records": "紀錄: {count}",
        "status.images": "圖片: {count}",
        "status.language": "語言: {language}",
        "table.index": "#",
        "table.date": "日期",
        "table.start": "上班",
        "table.end": "下班",
        "table.total": "總時數",
        "table.status": "狀態",
        "table.source": "來源",
        "selection.none": "未選擇",
        "selection.template": "Excel 範本",
        "selection.output": "輸出",
        "info.created": "建立時間:",
        "info.images": "圖片:",
        "info.records": "紀錄:",
        "info.updated": "最後更新:",
        "workflow.import.title": "匯入圖片",
        "workflow.import.body": "拖放或選擇打卡單照片",
        "workflow.recognize.title": "辨識",
        "workflow.recognize.body": "使用 OCR 擷取出勤資料",
        "workflow.review.title": "檢查與修正",
        "workflow.review.body": "確認並修正辨識結果",
        "workflow.export.title": "匯出",
        "workflow.export.body": "匯出時數表 Excel",
        "dialog.choose_photos": "選擇打卡單照片",
        "dialog.choose_template": "選擇 Excel 範本",
        "dialog.choose_output": "選擇輸出檔案",
        "dialog.excel_filter": "Excel (*.xlsx)",
        "dialog.image_filter": "圖片 (*.jpg *.jpeg *.png *.bmp *.webp)",
        "dialog.error": "錯誤",
        "dialog.done": "完成",
        "dialog.photo_limit_title": "照片數量限制",
        "dialog.photo_limit": "一次最多處理兩張照片，已先載入前兩張。",
        "error.no_photos": "請先選擇或拖入照片。",
        "error.too_many_photos": "一次最多處理兩張照片。",
        "error.photo_pair": "請提供一張 1-15 日照片，以及一張 16-31 日照片。",
        "error.bad_drop": "請拖入 jpg、jpeg、png、bmp 或 webp 圖片。",
        "error.no_template": "請先選擇 Excel 範本。",
        "error.invalid_rows": "以下日期有無效時間值: {rows}",
        "progress.loading_engine": "正在載入 OCR 引擎",
        "progress.engine_ready": "OCR 引擎已就緒",
        "progress.recognizing": "正在辨識 {name}",
        "progress.finished_photo": "已完成 {name}",
        "progress.done": "辨識完成",
        "status.photo_failed": "{name} 辨識失敗: {error}",
        "status.recognition_failed": "辨識失敗: {error}",
        "status.recognition_done": "辨識完成，請檢查表格後再匯出。",
        "status.learned_done": "辨識完成，已套用 {count} 筆修正紀錄，請檢查後再匯出。",
        "status.exported": "已匯出: {path}",
        "message.exported": "Excel 已匯出:\n{path}",
        "status.ok": "正常",
        "status.blank": "空白",
        "status.needs_review": "需確認",
        "status.failed": "失敗",
        "status.abnormal_shift": "異常班段",
        "status.short_shift": "短班提醒",
        "status.learned": "由修正紀錄套用",
        "tip.ok": "正常",
        "tip.blank": "未偵測到時間。",
        "tip.failed": "OCR 無法判斷此列。",
        "tip.shift": "班段長度短於或長於設定範圍。",
        "tip.review": "請檢查此列。",
        "filename.untitled": "未命名",
        "config.error": "設定錯誤",
    },
    "ja": {
        "app.title": "勤務時間集計ツール",
        "project.section": "プロジェクト",
        "workflow.section": "ワークフロー",
        "info.section": "プロジェクト情報",
        "field.name": "名前",
        "field.month": "月",
        "field.language": "言語",
        "placeholder.name": "従業員名またはプロジェクト名",
        "button.new_project": "新規",
        "button.files": "写真を選択",
        "button.template": "テンプレート",
        "button.output": "出力先",
        "button.recognize": "認識",
        "button.export": "Excel 出力",
        "button.settings": "設定",
        "drop.title": "ここに画像をドロップ",
        "drop.subtitle": "タイムカード画像をドラッグして開始",
        "drop.support": "JPG、JPEG、PNG、BMP、WEBP 対応",
        "drop.loaded": "読み込み済み画像",
        "status.ready": "準備完了",
        "status.records": "記録: {count}",
        "status.images": "画像: {count}",
        "status.language": "言語: {language}",
        "table.index": "#",
        "table.date": "日付",
        "table.start": "出勤",
        "table.end": "退勤",
        "table.total": "合計時間",
        "table.status": "状態",
        "table.source": "元画像",
        "selection.none": "未選択",
        "selection.template": "Excel テンプレート",
        "selection.output": "出力",
        "info.created": "作成:",
        "info.images": "画像:",
        "info.records": "記録:",
        "info.updated": "最終更新:",
        "workflow.import.title": "画像を読み込む",
        "workflow.import.body": "タイムカード画像をドラッグまたは選択",
        "workflow.recognize.title": "認識",
        "workflow.recognize.body": "OCR で勤怠データを抽出",
        "workflow.review.title": "確認と修正",
        "workflow.review.body": "認識結果を確認して修正",
        "workflow.export.title": "出力",
        "workflow.export.body": "勤怠データを Excel に出力",
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
        "error.bad_drop": "jpg、jpeg、png、bmp、webp 画像をドロップしてください。",
        "error.no_template": "先に Excel テンプレートを選択してください。",
        "error.invalid_rows": "無効な時間がある日付: {rows}",
        "progress.loading_engine": "OCR エンジンを読み込み中",
        "progress.engine_ready": "OCR エンジン準備完了",
        "progress.recognizing": "{name} を認識中",
        "progress.finished_photo": "{name} が完了しました",
        "progress.done": "認識完了",
        "status.photo_failed": "{name} の認識に失敗しました: {error}",
        "status.recognition_failed": "認識に失敗しました: {error}",
        "status.recognition_done": "認識が完了しました。出力前に表を確認してください。",
        "status.learned_done": "認識が完了しました。修正履歴を {count} 件適用しました。出力前に確認してください。",
        "status.exported": "出力しました: {path}",
        "message.exported": "Excel を出力しました:\n{path}",
        "status.ok": "正常",
        "status.blank": "空白",
        "status.needs_review": "要確認",
        "status.failed": "失敗",
        "status.abnormal_shift": "異常シフト",
        "status.short_shift": "短時間注意",
        "status.learned": "修正履歴から適用",
        "tip.ok": "正常",
        "tip.blank": "時間が検出されませんでした。",
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


def _format_duration(start_time: str, end_time: str) -> str:
    duration = _shift_minutes(start_time, end_time)
    if duration is None:
        return ""
    return f"{duration // 60}:{duration % 60:02d}"


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


def find_asset_path(filename: str) -> Optional[Path]:
    for base_dir in (RUNTIME_DIR, APP_DIR):
        candidate = base_dir / "assets" / filename
        if candidate.is_file():
            return candidate
    return None


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
            self._emit_progress(4, "loading_engine")
            engine = OCREngine()
            engine.initialize()
            self._emit_progress(12, "engine_ready")

            total = max(1, len(self.paths))
            for photo_index, path in enumerate(self.paths):
                self._emit_progress(12 + int(photo_index * 78 / total), "recognizing", path.name)
                try:
                    image = read_image(path)
                    fallback_half = self.config.classify_photo(path, photo_index)
                    recognized_rows = engine.recognize_attendance(image, fallback_half=fallback_half)
                except (ImageProcessingError, OCREngineError) as exc:
                    logging.exception("photo recognition failed")
                    fallback_half = self.config.classify_photo(path, photo_index)
                    self._mark_photo_failed(records, fallback_half, path.name)
                    self.progress.emit(f"photo_failed|{path.name}|{exc}")
                    self._emit_progress(12 + int((photo_index + 1) * 78 / total), "finished_photo", path.name)
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
                self._emit_progress(12 + int((photo_index + 1) * 78 / total), "finished_photo", path.name)

            self._emit_progress(100, "done")
            self.finished.emit([records[day] for day in range(1, 32)])
        except (ConfigError, OCREngineError) as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            logging.exception("recognition failed")
            self.failed.emit(str(exc))

    def _emit_progress(self, value: int, key: str, detail: str = "") -> None:
        value = max(0, min(100, int(value)))
        self.progress.emit(f"progress|{value}|{key}|{detail}")

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

        self.created_at = datetime.now()
        self.last_updated: Optional[datetime] = None
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
        self.workflow_cards: list[QFrame] = []
        self.workflow_titles: list[QLabel] = []
        self.workflow_bodies: list[QLabel] = []

        learning = self.config.learning
        self.correction_history = CorrectionHistory(
            RUNTIME_DIR / str(learning.get("history_file", "logs/correction_history.jsonl")),
            min_count=int(learning.get("min_count", 2)),
        )

        self.resize(1500, 940)
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
            QWidget#root { background: #f7f9fc; color: #172033; font-family: "Segoe UI", "Microsoft JhengHei", "Yu Gothic UI"; font-size: 14px; }
            QFrame#sidebar { background: #f8fbff; border-right: 1px solid #d6dde8; }
            QLabel#sectionLabel { color: #315b92; font-size: 12px; font-weight: 700; }
            QLabel#infoLabel { color: #4b5f79; }
            QLabel#infoValue { color: #172033; }
            QFrame#workflowCard { background: transparent; border: 1px solid transparent; border-radius: 6px; }
            QFrame#workflowCard[active="true"] { background: #edf6ff; border: 1px solid #94bdf5; }
            QLabel#stepBadge { background: #8a96a8; color: white; border-radius: 14px; min-width: 28px; max-width: 28px; min-height: 28px; max-height: 28px; font-weight: 700; }
            QLabel#stepBadge[active="true"] { background: #1d6bd1; }
            QLabel#workflowTitle { color: #172033; font-weight: 700; }
            QLabel#workflowBody { color: #4b5f79; }
            QLabel#dropZone { background: #fbfdff; border: 2px dashed #1f6fd3; border-radius: 9px; color: #23324a; padding: 26px; }
            QLabel#dropZone:hover { background: #f5faff; border-color: #0f5fc7; }
            QLabel#statusLabel { color: #314560; }
            QProgressBar { background: #edf2f8; border: 1px solid #cbd7e7; border-radius: 6px; height: 20px; text-align: center; color: #172033; }
            QProgressBar::chunk { background: #1f6fd3; border-radius: 5px; }
            QTableWidget { background: #ffffff; alternate-background-color: #f9fbff; color: #172033; gridline-color: #e1e7f0; border: 1px solid #d9e1ec; border-radius: 7px; }
            QTableWidget::item { padding: 4px; }
            QTableWidget::item:selected { background: #dbeafe; color: #172033; }
            QHeaderView::section { background: #f4f7fb; border: 0; border-right: 1px solid #d8e1ed; border-bottom: 1px solid #d8e1ed; color: #34445c; font-weight: 700; padding: 7px; }
            QPushButton { padding: 8px 12px; border-radius: 6px; border: 1px solid #cbd7e6; background: #ffffff; color: #1f2b3d; }
            QPushButton:hover { border-color: #8eb6ef; background: #f6faff; }
            QPushButton:disabled { color: #8090a4; background: #eef2f7; }
            QPushButton#primaryButton { background: #1769d2; border-color: #1769d2; color: white; font-weight: 700; }
            QPushButton#primaryButton:hover { background: #0f5fc7; }
            QPushButton#secondaryBlueButton { background: #ffffff; border-color: #bed0e9; color: #1769d2; font-weight: 700; }
            QLineEdit, QSpinBox, QComboBox { padding: 7px; border: 1px solid #cbd5e1; border-radius: 5px; background: white; color: #172033; }
            QComboBox QAbstractItemView { background: #ffffff; color: #172033; border: 1px solid #cbd5e1; selection-background-color: #dbeafe; selection-color: #172033; outline: 0; }
            QFrame#bottomBar { background: #f8fbff; border-top: 1px solid #d6dde8; }
            """
        )

        shell = QVBoxLayout(root)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)

        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)
        shell.addLayout(body, 1)

        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(335)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(18, 24, 18, 18)
        sidebar_layout.setSpacing(16)
        body.addWidget(sidebar)

        self.project_section_label = self._section_label()
        sidebar_layout.addWidget(self.project_section_label)
        self.name_input = QLineEdit()
        self.month_input = QSpinBox()
        self.month_input.setRange(1, 12)
        self.month_input.setValue(get_default_month())

        project_grid = QGridLayout()
        project_grid.setHorizontalSpacing(8)
        project_grid.setVerticalSpacing(8)
        self.name_label = QLabel()
        self.month_label = QLabel()
        project_grid.addWidget(self.name_label, 0, 0)
        project_grid.addWidget(self.name_input, 0, 1)
        project_grid.addWidget(self.month_label, 1, 0)
        project_grid.addWidget(self.month_input, 1, 1)
        sidebar_layout.addLayout(project_grid)

        project_buttons = QHBoxLayout()
        self.new_project_button = QPushButton()
        self.files_button = QPushButton()
        self.new_project_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        self.files_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
        project_buttons.addWidget(self.new_project_button)
        project_buttons.addWidget(self.files_button)
        sidebar_layout.addLayout(project_buttons)

        self.workflow_section_label = self._section_label()
        sidebar_layout.addWidget(self.workflow_section_label)
        self._add_workflow(sidebar_layout, 1, "workflow.import.title", "workflow.import.body")
        self._add_workflow(sidebar_layout, 2, "workflow.recognize.title", "workflow.recognize.body")
        self._add_workflow(sidebar_layout, 3, "workflow.review.title", "workflow.review.body")
        self._add_workflow(sidebar_layout, 4, "workflow.export.title", "workflow.export.body")
        sidebar_layout.addStretch(1)

        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setStyleSheet("color: #d6dde8;")
        sidebar_layout.addWidget(divider)

        self.info_section_label = self._section_label()
        sidebar_layout.addWidget(self.info_section_label)
        self.info_created_value = self._info_row(sidebar_layout)
        self.info_images_value = self._info_row(sidebar_layout)
        self.info_records_value = self._info_row(sidebar_layout)
        self.info_updated_value = self._info_row(sidebar_layout)

        settings_buttons = QHBoxLayout()
        self.template_button = QPushButton()
        self.output_button = QPushButton()
        self.template_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        self.output_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        settings_buttons.addWidget(self.template_button)
        settings_buttons.addWidget(self.output_button)
        sidebar_layout.addLayout(settings_buttons)

        main_area = QWidget()
        main_area_layout = QVBoxLayout(main_area)
        main_area_layout.setContentsMargins(18, 18, 18, 16)
        main_area_layout.setSpacing(14)
        body.addWidget(main_area, 1)

        toolbar = QHBoxLayout()
        toolbar.setSpacing(10)
        self.language_label = QLabel()
        self.language_combo = QComboBox()
        self.language_combo.setFixedWidth(150)
        self.language_combo.view().setStyleSheet(
            "QAbstractItemView { background: #ffffff; color: #172033; "
            "selection-background-color: #dbeafe; selection-color: #172033; "
            "border: 1px solid #cbd5e1; outline: 0; }"
        )
        for code, label in LANGUAGES:
            self.language_combo.addItem(label, code)
        index = self.language_combo.findData(self.language)
        self.language_combo.setCurrentIndex(max(index, 0))
        self.language_combo.currentIndexChanged.connect(self._language_changed)
        toolbar.addWidget(self.language_label)
        toolbar.addWidget(self.language_combo)
        toolbar.addStretch(1)

        self.recognize_button = QPushButton()
        self.recognize_button.setObjectName("secondaryBlueButton")
        self.recognize_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        self.export_button = QPushButton()
        self.export_button.setObjectName("primaryButton")
        self.export_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        toolbar.addWidget(self.recognize_button)
        toolbar.addWidget(self.export_button)
        main_area_layout.addLayout(toolbar)

        if self.config.ui.get("enable_drag_drop", True):
            self.selection_label = DropImageLabel()
            self.selection_label.filesDropped.connect(self._set_photo_paths_from_drop)
        else:
            self.selection_label = QLabel()
        self.selection_label.setObjectName("dropZone")
        self.selection_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.selection_label.setWordWrap(True)
        self.selection_label.setTextFormat(Qt.TextFormat.RichText)
        self.selection_label.setMinimumHeight(175)
        self.selection_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.selection_label.setStyleSheet("")
        main_area_layout.addWidget(self.selection_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        main_area_layout.addWidget(self.progress_bar)

        self.table = QTableWidget(31, 7)
        self.table.setAlternatingRowColors(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(25)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked | QTableWidget.EditTrigger.SelectedClicked | QTableWidget.EditTrigger.EditKeyPressed)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(COL_INDEX, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(COL_TOTAL, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(COL_STATUS, QHeaderView.ResizeMode.ResizeToContents)
        main_area_layout.addWidget(self.table, 1)

        bottom_bar = QFrame()
        bottom_bar.setObjectName("bottomBar")
        bottom_layout = QHBoxLayout(bottom_bar)
        bottom_layout.setContentsMargins(18, 10, 18, 10)
        bottom_layout.setSpacing(18)
        self.bottom_images_label = QLabel()
        self.bottom_records_label = QLabel()
        self.status_label = QLabel()
        self.status_label.setObjectName("statusLabel")
        self.bottom_language_label = QLabel()
        self.bottom_time_label = QLabel(datetime.now().strftime("%H:%M"))
        for label in (self.bottom_images_label, self.bottom_records_label, self.status_label, self.bottom_language_label, self.bottom_time_label):
            bottom_layout.addWidget(label)
        bottom_layout.addStretch(1)
        shell.addWidget(bottom_bar)

        self.setCentralWidget(root)
        self.new_project_button.clicked.connect(self._reset_project)
        self.files_button.clicked.connect(self._select_files)
        self.template_button.clicked.connect(self._select_template)
        self.output_button.clicked.connect(self._select_output)
        self.recognize_button.clicked.connect(self._start_recognition)
        self.export_button.clicked.connect(self._export)
        self.table.cellChanged.connect(self._cell_changed)

    def _section_label(self) -> QLabel:
        label = QLabel()
        label.setObjectName("sectionLabel")
        return label

    def _add_workflow(self, parent: QVBoxLayout, number: int, title_key: str, body_key: str) -> None:
        card = QFrame()
        card.setObjectName("workflowCard")
        card.setProperty("active", False)
        layout = QHBoxLayout(card)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)
        badge = QLabel(str(number))
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        badge.setObjectName("stepBadge")
        badge.setProperty("active", False)
        text_box = QVBoxLayout()
        text_box.setSpacing(3)
        title = QLabel(title_key)
        title.setObjectName("workflowTitle")
        body = QLabel(body_key)
        body.setObjectName("workflowBody")
        body.setWordWrap(True)
        text_box.addWidget(title)
        text_box.addWidget(body)
        layout.addWidget(badge)
        layout.addLayout(text_box, 1)
        parent.addWidget(card)
        self.workflow_cards.append(card)
        self.workflow_titles.append(title)
        self.workflow_bodies.append(body)

    def _info_row(self, parent: QVBoxLayout) -> QLabel:
        row = QHBoxLayout()
        title = QLabel()
        title.setObjectName("infoLabel")
        value = QLabel()
        value.setObjectName("infoValue")
        value.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        row.addWidget(title)
        row.addWidget(value, 1)
        parent.addLayout(row)
        value.setProperty("titleLabel", title)
        return value

    def _apply_language(self) -> None:
        self.setWindowTitle(self.tr("app.title"))
        self.project_section_label.setText(self.tr("project.section"))
        self.workflow_section_label.setText(self.tr("workflow.section"))
        self.info_section_label.setText(self.tr("info.section"))
        self.name_label.setText(self.tr("field.name"))
        self.month_label.setText(self.tr("field.month"))
        self.language_label.setText(self.tr("field.language"))
        self.name_input.setPlaceholderText(self.tr("placeholder.name"))
        self.new_project_button.setText(self.tr("button.new_project"))
        self.files_button.setText(self.tr("button.files"))
        self.template_button.setText(self.tr("button.template"))
        self.output_button.setText(self.tr("button.output"))
        self.recognize_button.setText(self.tr("button.recognize"))
        self.export_button.setText(self.tr("button.export"))
        for title, key in zip(
            self.workflow_titles,
            ("workflow.import.title", "workflow.recognize.title", "workflow.review.title", "workflow.export.title"),
        ):
            title.setText(self.tr(key))
        for body, key in zip(
            self.workflow_bodies,
            ("workflow.import.body", "workflow.recognize.body", "workflow.review.body", "workflow.export.body"),
        ):
            body.setText(self.tr(key))
        for value, key in (
            (self.info_created_value, "info.created"),
            (self.info_images_value, "info.images"),
            (self.info_records_value, "info.records"),
            (self.info_updated_value, "info.updated"),
        ):
            title = value.property("titleLabel")
            if isinstance(title, QLabel):
                title.setText(self.tr(key))
        self.table.setHorizontalHeaderLabels(
            (
                self.tr("table.index"),
                self.tr("table.date"),
                self.tr("table.start"),
                self.tr("table.end"),
                self.tr("table.total"),
                self.tr("table.status"),
                self.tr("table.source"),
            )
        )
        if not self.status_label.text():
            self.status_label.setText(self.tr("status.ready"))
        self._update_selection_label()
        self._update_project_info()

    def _language_changed(self, *_args: object) -> None:
        code = self.language_combo.currentData()
        if code not in TRANSLATIONS:
            return
        self.language = code
        self.settings.setValue("language", self.language)
        self._apply_language()
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
        self.last_updated = datetime.now()
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
        if not hasattr(self, "selection_label"):
            return
        template = self.template_path.name if self.template_path else self.tr("selection.none")
        output_path = self._current_output_path()
        icon_path = find_asset_path("drop-image.svg")
        icon_html = (
            f"<img src='{icon_path.resolve().as_uri()}' width='64' height='64'/>"
            if icon_path
            else "<div style='font-size:42px;color:#1f6fd3;'>□</div>"
        )
        if self.photo_paths:
            files = "".join(f"<div style='margin-top:4px;color:#1f3b5d;'>{index}. {path.name}</div>" for index, path in enumerate(self.photo_paths, 1))
            body = (
                f"{icon_html}"
                f"<div style='font-size:18px;font-weight:700;'>{self.tr('drop.loaded')}</div>"
                f"{files}"
            )
        else:
            body = (
                f"{icon_html}"
                f"<div style='font-size:18px;font-weight:700;'>{self.tr('drop.title')}</div>"
                f"<div style='margin-top:8px;color:#4b5f79;'>{self.tr('drop.subtitle')}</div>"
                f"<div style='margin-top:8px;color:#4b5f79;'>{self.tr('drop.support')}</div>"
            )
        footer = (
            f"<div style='margin-top:12px;color:#60718a;font-size:12px;'>"
            f"{self.tr('selection.template')}: {template} &nbsp;&nbsp; "
            f"{self.tr('selection.output')}: {output_path}"
            f"</div>"
        )
        self.selection_label.setText(f"<div align='center'>{body}{footer}</div>")
        self._update_project_info()
        self._update_workflow_state()

    def _update_project_info(self) -> None:
        if not hasattr(self, "info_created_value"):
            return
        record_count = sum(1 for record in self.records if record.status != STATUS_BLANK or record.start_time or record.end_time)
        self.info_created_value.setText(self.created_at.strftime("%Y-%m-%d %H:%M"))
        self.info_images_value.setText(str(len(self.photo_paths)))
        self.info_records_value.setText(str(record_count))
        self.info_updated_value.setText(self.last_updated.strftime("%Y-%m-%d %H:%M") if self.last_updated else "-")
        self.bottom_images_label.setText(self.tr("status.images", count=len(self.photo_paths)))
        self.bottom_records_label.setText(self.tr("status.records", count=record_count))
        self.bottom_language_label.setText(self.tr("status.language", language=self.language_combo.currentText()))

    def _update_workflow_state(self) -> None:
        active_index = 0
        if self.photo_paths:
            active_index = 1
        if any(record.status != STATUS_BLANK or record.start_time or record.end_time for record in self.records):
            active_index = 2
        if self.output_path:
            active_index = 3
        for index, card in enumerate(self.workflow_cards):
            is_active = index == active_index
            card.setProperty("active", is_active)
            for child in card.findChildren(QLabel):
                if child.objectName() == "stepBadge":
                    child.setProperty("active", is_active)
                    child.style().unpolish(child)
                    child.style().polish(child)
            card.style().unpolish(card)
            card.style().polish(card)

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

    def _reset_project(self) -> None:
        self.photo_paths = []
        self.output_path = None
        self._manual_output_path = False
        self.records = [AttendanceRecord(day=day) for day in range(1, 32)]
        self.last_updated = datetime.now()
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)
        self.status_label.setText(self.tr("status.ready"))
        self._update_selection_label()
        self._refresh_table()

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
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat(self.tr("progress.loading_engine"))
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
        parts = payload.split("|", 3)
        if parts[0] == "progress" and len(parts) >= 3:
            try:
                value = int(parts[1])
            except ValueError:
                value = self.progress_bar.value()
            key = parts[2]
            detail = parts[3] if len(parts) > 3 else ""
            message = self.tr(f"progress.{key}", name=detail)
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(value)
            self.progress_bar.setFormat(message)
            self.status_label.setText(message)
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
        self.last_updated = datetime.now()
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat(self.tr("progress.done"))
        self._refresh_table()
        if learned_count:
            self.status_label.setText(self.tr("status.learned_done", count=learned_count))
        else:
            self.status_label.setText(self.tr("status.recognition_done"))

    def _recognition_failed(self, message: str) -> None:
        self.status_label.setText(self.tr("status.recognition_failed", error=message))
        self.progress_bar.setVisible(False)
        self._show_error(self.tr("status.recognition_failed", error=message))

    def _thread_finished(self) -> None:
        self._set_busy(False)
        if self.worker:
            self.worker.deleteLater()
        self.worker = None
        self.thread = None

    def _set_busy(self, busy: bool) -> None:
        for button in (
            self.new_project_button,
            self.files_button,
            self.template_button,
            self.output_button,
            self.recognize_button,
            self.export_button,
        ):
            button.setEnabled(not busy)
        self.language_combo.setEnabled(not busy)

    def _date_label(self, day: int) -> str:
        return f"{self.month_input.value():02d}/{day:02d}"

    def _display_status(self, record: AttendanceRecord) -> str:
        status = self.tr(f"status.{record.status}")
        if record.learned_fields:
            return f"{status} / {self.tr('status.learned')}"
        return status

    def _refresh_table(self) -> None:
        self._updating_table = True
        try:
            for row, record in enumerate(self.records):
                values = [
                    str(row + 1),
                    self._date_label(record.day),
                    record.start_time or "",
                    record.end_time or "",
                    _format_duration(record.start_time or "", record.end_time or ""),
                    self._display_status(record),
                    Path(record.source_image).name if record.source_image else "",
                ]
                for column, value in enumerate(values):
                    item = QTableWidgetItem(value)
                    if column in (COL_INDEX, COL_DATE, COL_TOTAL, COL_STATUS, COL_SOURCE):
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    if column in (COL_INDEX, COL_TOTAL, COL_STATUS):
                        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.table.setItem(row, column, item)
                    self._last_table_values[(row, column)] = value
                self._apply_row_style(row, record)
        finally:
            self._updating_table = False
        self._update_project_info()

    def _apply_row_style(self, row: int, record: AttendanceRecord) -> None:
        if not self.config.ui.get("highlight_abnormal_rows", True):
            return
        status = record.status or ""
        learned = bool(record.learned_fields)
        status_bg = QColor(224, 242, 254) if learned else QColor(234, 247, 233)
        if status == STATUS_OK:
            bg = QColor(224, 242, 254) if learned else QColor(255, 255, 255)
            tip = self.tr("status.learned") if learned else self.tr("tip.ok")
        elif status == STATUS_BLANK:
            bg = QColor(246, 247, 249)
            status_bg = QColor(236, 239, 244)
            tip = self.tr("tip.blank")
        elif status == STATUS_FAILED:
            bg = QColor(255, 226, 226)
            status_bg = QColor(255, 204, 204)
            tip = self.tr("tip.failed")
        elif status in {STATUS_ABNORMAL, STATUS_SHORT}:
            bg = QColor(255, 245, 225)
            status_bg = QColor(255, 229, 180)
            tip = self.tr("tip.shift")
        else:
            bg = QColor(255, 248, 218)
            status_bg = QColor(255, 238, 169)
            tip = self.tr("tip.review")
        for column in range(self.table.columnCount()):
            item = self.table.item(row, column)
            if item is not None:
                item.setBackground(QBrush(status_bg if column == COL_STATUS else bg))
                item.setForeground(QBrush(QColor(23, 32, 51)))
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
        self.last_updated = datetime.now()
        self._updating_table = True
        try:
            if item and new_value and item.text() != new_value:
                item.setText(new_value)
            total_item = self.table.item(row, COL_TOTAL)
            if total_item:
                total_item.setText(_format_duration(record.start_time or "", record.end_time or ""))
            status_item = self.table.item(row, COL_STATUS)
            if status_item:
                status_item.setText(self._display_status(record))
        finally:
            self._updating_table = False
        self._apply_row_style(row, record)
        self._update_project_info()

    def _export(self) -> None:
        if not self.template_path:
            self._show_error(self.tr("error.no_template"))
            return
        invalid_rows = []
        for row in range(self.table.rowCount()):
            for column in (COL_START, COL_END):
                cell = self.table.item(row, column)
                text = cell.text().strip() if cell else ""
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
            self.output_path = output_path
            self._manual_output_path = True
            self.last_updated = datetime.now()
            self.settings.setValue("last_output_dir", str(self.output_dir))
            self.status_label.setText(self.tr("status.exported", path=output_path))
            self._update_selection_label()
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
