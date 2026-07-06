from __future__ import annotations

from collections import defaultdict
from typing import List, Tuple
import logging
import re

import cv2
import numpy as np

from models import OCRAttendanceRow, OCRResult
from utils import extract_ocr_time_text, extract_time_text


class OCREngineError(RuntimeError):
    pass


class OCREngine:
    def __init__(self) -> None:
        self._ocr = None

    def initialize(self) -> None:
        if self._ocr is not None:
            return
        try:
            from rapidocr import RapidOCR

            self._ocr = RapidOCR()
        except Exception as exc:
            logging.exception("RapidOCR initialization failed")
            raise OCREngineError(f"OCR initialization failed: {exc}") from exc

    def recognize(self, image: np.ndarray) -> OCRResult:
        self.initialize()
        try:
            raw_result = self._ocr(image)
            texts = getattr(raw_result, "txts", None)
            scores = getattr(raw_result, "scores", None)
            texts = [] if texts is None else list(texts)
            scores = [] if scores is None else list(scores)
            items = [OCRResult(str(text), float(score)) for text, score in zip(texts, scores)]
            return max(items, key=lambda item: item.confidence) if items else OCRResult()
        except Exception as exc:
            logging.exception("ROI OCR failed")
            raise OCREngineError(f"ROI OCR failed: {exc}") from exc

    def recognize_attendance(self, image: np.ndarray, fallback_half: str) -> List[OCRAttendanceRow]:
        self.initialize()
        try:
            enhanced = self._enhance(image)
            detections = self._detect(enhanced)
            header = self._find_header(detections, image.shape[0])
            detected_days = self._detect_days(detections, header_bottom=header[3] if header else 0)
            if header is None:
                header = self._infer_header_from_days(detections, image.shape[1], image.shape[0])
            if header is None:
                raise OCREngineError("Could not find the attendance table header. Please retake the photo so the date column is clear.")

            left, _, right, header_bottom = header
            table_width = max(1.0, right - left)
            if not detected_days:
                detected_days = self._detect_days(detections, header_bottom=header_bottom)

            half = self._classify_half([day for day, _ in detected_days]) or fallback_half
            first_day, row_count = (1, 15) if half == "first_half" else (16, 16)
            day_points = [(day, y) for day, y in detected_days if first_day <= day < first_day + row_count]
            row_origin, row_step = self._fit_rows(day_points, header_bottom, row_count)

            grouped: dict[int, dict[str, list[tuple[str, float]]]] = defaultdict(lambda: defaultdict(list))
            for bbox, text, score in detections:
                if not self._looks_like_time_fragment(text):
                    continue
                center_x, center_y = self._center(bbox)
                if center_y <= header_bottom:
                    continue
                row_index = round((center_y - row_origin) / row_step)
                day = first_day + row_index
                if day < first_day or day >= first_day + row_count:
                    continue
                relative_x = (center_x - left) / table_width
                if 0.43 <= relative_x <= 0.67:
                    column = "end"
                elif 0.55 <= relative_x <= 0.92:
                    column = "start"
                else:
                    continue
                grouped[day][column].append((text, score))

            rows = []
            for day in range(first_day, first_day + row_count):
                start_text, start_score = self._choose_time(grouped[day]["start"])
                end_text, end_score = self._choose_time(grouped[day]["end"])
                rows.append(
                    OCRAttendanceRow(
                        day=day,
                        start_text=start_text,
                        end_text=end_text,
                        confidence=min(start_score, end_score) if start_text and end_text else max(start_score, end_score),
                    )
                )
            return rows
        except OCREngineError:
            raise
        except Exception as exc:
            logging.exception("attendance OCR failed")
            raise OCREngineError(f"Attendance OCR failed: {exc}") from exc

    def _detect(self, image: np.ndarray) -> List[Tuple[Tuple[float, float, float, float], str, float]]:
        result = self._ocr(image)
        boxes = getattr(result, "boxes", None)
        texts = getattr(result, "txts", None)
        scores = getattr(result, "scores", None)
        boxes = [] if boxes is None else list(boxes)
        texts = [] if texts is None else list(texts)
        scores = [] if scores is None else list(scores)
        return [(self._bbox(box), str(text).strip(), float(score)) for box, text, score in zip(boxes, texts, scores)]

    @staticmethod
    def _enhance(image: np.ndarray) -> np.ndarray:
        lab = cv2.cvtColor(image, cv2.COLOR_BGR2LAB)
        lightness, channel_a, channel_b = cv2.split(lab)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(lightness)
        return cv2.cvtColor(cv2.merge((enhanced, channel_a, channel_b)), cv2.COLOR_LAB2BGR)

    @staticmethod
    def _bbox(box) -> Tuple[float, float, float, float]:
        xs = [float(point[0]) for point in box]
        ys = [float(point[1]) for point in box]
        return min(xs), min(ys), max(xs), max(ys)

    @staticmethod
    def _center(bbox: Tuple[float, float, float, float]) -> Tuple[float, float]:
        left, top, right, bottom = bbox
        return (left + right) / 2, (top + bottom) / 2

    @staticmethod
    def _detect_days(detections, header_bottom: float = 0) -> list[tuple[int, float]]:
        detected_days = []
        for bbox, text, _score in detections:
            center_x, center_y = OCREngine._center(bbox)
            compact = text.strip()
            if center_y <= header_bottom or not re.fullmatch(r"\d{1,2}", compact):
                continue
            day = int(compact)
            if 1 <= day <= 31:
                detected_days.append((day, center_y))
        return detected_days

    @staticmethod
    def _find_header(detections, image_height: int):
        keywords = ("日期", "上班", "下班", "結束", "加班", "備註", "date", "time in", "time out")
        candidates = []
        for bbox, text, _ in detections:
            _, center_y = OCREngine._center(bbox)
            if center_y > image_height * 0.55:
                continue
            lowered = text.lower()
            if any(keyword in text or keyword in lowered for keyword in keywords):
                candidates.append((bbox, text))
        if not candidates:
            return None
        anchor_bbox, _ = min(candidates, key=lambda item: OCREngine._center(item[0])[1])
        anchor_y = OCREngine._center(anchor_bbox)[1]
        y_tolerance = max(35.0, image_height * 0.025)
        row_boxes = [bbox for bbox, _ in candidates if abs(OCREngine._center(bbox)[1] - anchor_y) <= y_tolerance]
        if not row_boxes:
            row_boxes = [anchor_bbox]
        return (
            min(bbox[0] for bbox in row_boxes),
            min(bbox[1] for bbox in row_boxes),
            max(bbox[2] for bbox in row_boxes),
            max(bbox[3] for bbox in row_boxes),
        )

    @staticmethod
    def _infer_header_from_days(detections, image_width: int, image_height: int):
        days = []
        for bbox, text, _score in detections:
            compact = text.strip()
            if not re.fullmatch(r"\d{1,2}", compact):
                continue
            day = int(compact)
            if 1 <= day <= 31:
                center_x, center_y = OCREngine._center(bbox)
                if center_y < image_height * 0.92:
                    days.append((day, center_x, center_y, bbox))
        if len(days) < 3:
            return None
        y_values = sorted(center_y for _day, _x, center_y, _bbox in days)
        steps = [b - a for a, b in zip(y_values, y_values[1:]) if 4 <= b - a <= image_height * 0.12]
        row_step = float(np.median(steps)) if steps else max(18.0, image_height * 0.035)
        first_row_y = min(center_y for _day, _x, center_y, _bbox in days)
        left = max(0.0, min(bbox[0] for _day, _x, _y, bbox in days) - image_width * 0.02)
        right = image_width * 0.95
        header_bottom = max(0.0, first_row_y - row_step * 0.55)
        return (left, max(0.0, header_bottom - row_step), right, header_bottom)

    @staticmethod
    def _fit_rows(day_points, header_bottom: float, row_count: int) -> Tuple[float, float]:
        first_expected_day = 1 if row_count == 15 else 16
        if len(day_points) >= 2:
            sorted_points = sorted(day_points)
            steps = []
            for (day_a, y_a), (day_b, y_b) in zip(sorted_points, sorted_points[1:]):
                gap = day_b - day_a
                if gap > 0:
                    steps.append((y_b - y_a) / gap)
            if steps:
                row_step = float(np.median(steps))
                first_day, first_y = sorted_points[0]
                row_origin = first_y - (first_day - first_expected_day) * row_step
                return row_origin, row_step
        return header_bottom + 20, max(8.0, (1000 - header_bottom) / row_count)

    @staticmethod
    def _looks_like_time_fragment(text: str) -> bool:
        compact = str(text or "").replace(" ", "")
        return bool(extract_ocr_time_text(compact) or re.fullmatch(r"\d{1,4}", compact))

    @staticmethod
    def _choose_time(items: list[tuple[str, float]]) -> tuple[str, float]:
        candidates = []
        for text, score in items:
            normalized = extract_ocr_time_text(text) or extract_time_text(text)
            if normalized:
                candidates.append((normalized, text, score))
        if not candidates:
            return "", 0.0
        best = max(candidates, key=lambda item: (item[2], len(item[1])))
        return best[0], float(best[2])

    @staticmethod
    def _classify_half(detected_days: list[int]) -> str | None:
        first = sum(1 for day in detected_days if day <= 15)
        second = sum(1 for day in detected_days if day >= 16)
        if first == second:
            return None
        return "first_half" if first > second else "second_half"
