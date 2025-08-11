"""OCR engine using pytesseract for screen text detection."""

from __future__ import annotations

import time
from typing import Optional, Tuple

import pytesseract
from PIL import ImageGrab
import pyautogui

from .config import OCR_CONFIDENCE, OCR_LANGUAGE
from .logger import get_logger

logger = get_logger(__name__)


class OCREngine:
    def __init__(self):
        self.lang = OCR_LANGUAGE

    def _screenshot(self, region=None):
        try:
            img = ImageGrab.grab(bbox=region)
            return img
        except Exception as exc:
            logger.error("Screenshot failed: %s", exc)
            return None

    def find_text_on_screen(self, text: str, region=None, confidence: float = OCR_CONFIDENCE) -> Optional[Tuple[int, int, int, int]]:
        """Find text coordinates using OCR."""
        img = self._screenshot(region)
        if img is None:
            return None
        data = pytesseract.image_to_data(img, lang=self.lang, output_type=pytesseract.Output.DICT)
        for i, found_text in enumerate(data["text"]):
            if found_text.strip().lower() == text.lower() and float(data["conf"][i]) / 100 >= confidence:
                x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
                if region:
                    x += region[0]
                    y += region[1]
                return x, y, w, h
        return None

    def click_text(self, text: str, offset_x: int = 0, offset_y: int = 0) -> bool:
        """Click on found text with optional offset."""
        coords = self.find_text_on_screen(text)
        if not coords:
            logger.error("Text '%s' not found on screen", text)
            return False
        x, y, w, h = coords
        pyautogui.click(x + w // 2 + offset_x, y + h // 2 + offset_y)
        time.sleep(0.1)
        return True

    def wait_for_text(self, text: str, timeout: float = 10) -> bool:
        """Wait until text appears on screen."""
        end_time = time.time() + timeout
        while time.time() < end_time:
            if self.find_text_on_screen(text):
                return True
            time.sleep(0.5)
        logger.error("Timeout waiting for text: %s", text)
        return False
