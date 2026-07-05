from __future__ import annotations

from pathlib import Path
from typing import Tuple
import cv2
import numpy as np


class ImageProcessingError(RuntimeError):
    pass


def read_image(path: str | Path) -> np.ndarray:
    try:
        data = np.fromfile(str(path), dtype=np.uint8)
        image = cv2.imdecode(data, cv2.IMREAD_COLOR)
    except OSError as exc:
        raise ImageProcessingError(f"無法讀取圖片: {Path(path).name}") from exc
    if image is None:
        raise ImageProcessingError(f"不支援或損壞的圖片: {Path(path).name}")
    return image


def _rotate(image: np.ndarray, angle: float) -> np.ndarray:
    height, width = image.shape[:2]
    matrix = cv2.getRotationMatrix2D((width / 2, height / 2), angle, 1.0)
    return cv2.warpAffine(image, matrix, (width, height), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)


def deskew(image: np.ndarray) -> Tuple[np.ndarray, float]:
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 50, 150)
    height, width = gray.shape[:2]
    lines = cv2.HoughLinesP(edges, 1, np.pi / 180, max(80, width // 8), minLineLength=width // 4, maxLineGap=20)
    if lines is None:
        return image, 0.0
    angles = []
    for x1, y1, x2, y2 in lines[:, 0]:
        angle = np.degrees(np.arctan2(y2 - y1, x2 - x1))
        if abs(angle) < 15:
            angles.append(float(angle))
    if not angles:
        return image, 0.0
    correction = float(np.median(angles))
    return _rotate(image, correction), correction


def _order_points(points: np.ndarray) -> np.ndarray:
    points = points.reshape(4, 2)
    ordered = np.zeros((4, 2), dtype="float32")
    sums = points.sum(axis=1)
    differences = np.diff(points, axis=1)
    ordered[0] = points[np.argmin(sums)]
    ordered[2] = points[np.argmax(sums)]
    ordered[1] = points[np.argmin(differences)]
    ordered[3] = points[np.argmax(differences)]
    return ordered


def perspective_correct(image: np.ndarray) -> np.ndarray:
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blurred, 40, 140)
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    image_area = image.shape[0] * image.shape[1]
    for contour in sorted(contours, key=cv2.contourArea, reverse=True)[:8]:
        if cv2.contourArea(contour) < image_area * 0.25:
            continue
        perimeter = cv2.arcLength(contour, True)
        polygon = cv2.approxPolyDP(contour, 0.02 * perimeter, True)
        if len(polygon) != 4:
            continue
        top_left, top_right, bottom_right, bottom_left = _order_points(polygon)
        width = int(max(np.linalg.norm(bottom_right - bottom_left), np.linalg.norm(top_right - top_left)))
        height = int(max(np.linalg.norm(top_right - bottom_right), np.linalg.norm(top_left - bottom_left)))
        if width <= 0 or height <= 0:
            continue
        destination = np.array([[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]], dtype="float32")
        matrix = cv2.getPerspectiveTransform(np.array([top_left, top_right, bottom_right, bottom_left]), destination)
        return cv2.warpPerspective(image, matrix, (width, height))
    return image


def crop_relative(image: np.ndarray, bbox: list[float] | tuple[float, float, float, float]) -> np.ndarray:
    x1, y1, x2, y2 = [float(value) for value in bbox]
    height, width = image.shape[:2]
    left = max(0, min(width, int(x1 * width)))
    right = max(0, min(width, int(x2 * width)))
    top = max(0, min(height, int(y1 * height)))
    bottom = max(0, min(height, int(y2 * height)))
    if right <= left or bottom <= top:
        raise ImageProcessingError("ROI 範圍錯誤")
    return image[top:bottom, left:right]


def preprocess_photo(path: str | Path, table_bbox: list[float]) -> np.ndarray:
    image = read_image(path)
    corrected, _ = deskew(image)
    corrected = perspective_correct(corrected)
    return crop_relative(corrected, table_bbox)


def classify_half_by_color(image: np.ndarray) -> str:
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    hue = hsv[:, :, 0]
    saturation = hsv[:, :, 1]
    colored = saturation > 70
    red_count = int(np.count_nonzero(((hue < 12) | (hue > 165)) & colored))
    green_count = int(np.count_nonzero(((hue > 35) & (hue < 90)) & colored))
    return "first_half" if red_count >= green_count else "second_half"


def preprocess_cell(image: np.ndarray) -> np.ndarray:
    if image.size == 0:
        raise ImageProcessingError("空白儲存格影像")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if image.ndim == 3 else image
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)
    denoised = cv2.fastNlMeansDenoising(enhanced)
    _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    return cv2.copyMakeBorder(binary, 10, 10, 10, 10, cv2.BORDER_CONSTANT, value=255)
