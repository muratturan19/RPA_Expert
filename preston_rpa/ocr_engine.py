"""OCR engine using pytesseract for screen text detection."""

from __future__ import annotations

import time
from datetime import datetime
from typing import Optional, Tuple
from pathlib import Path

import pytesseract
import pyautogui
import pygetwindow as gw

from .config import OCR_CONFIDENCE, OCR_LANGUAGE, OCR_TESSERACT_CONFIG
from .logger import get_logger

logger = get_logger(__name__)


class OCREngine:
    def __init__(self, debug: bool = False):
        self.lang = OCR_LANGUAGE
        self.debug = debug

    def _save_debug_image(self, img, name: str) -> None:
        """Save screenshot for debugging purposes."""
        try:
            debug_dir = Path("debug")
            debug_dir.mkdir(exist_ok=True)
            img.save(debug_dir / f"{name}.png")
        except Exception as exc:
            logger.error("Failed to save debug image: %s", exc)

    def _screenshot(self, region=None):
        try:
            windows = gw.getWindowsWithTitle("Preston")
            if windows:
                windows[0].activate()
                time.sleep(2)
            img = pyautogui.screenshot(region=region)
            if self.debug:
                debug_dir = Path("debug")
                debug_dir.mkdir(exist_ok=True)
                filename = debug_dir / f"debug_{datetime.now().strftime('%H%M%S')}.png"
                img.save(filename)
                print("Screenshot saved, looking for text...")
                text = pytesseract.image_to_string(img, lang="tur")
                print(f"ALL OCR TEXT FOUND: {text}")
            return img
        except Exception as exc:
            logger.error("Screenshot failed: %s", exc)
            return None

    def find_text_on_screen(
        self,
        text: str,
        region=None,
        confidence: float = OCR_CONFIDENCE,
    ) -> Optional[Tuple[int, int, int, int]]:
        """Find text coordinates using OCR."""
        img = self._screenshot(region)
        if img is None:
            return None
        data = pytesseract.image_to_data(
            img,
            lang=self.lang,
            config=OCR_TESSERACT_CONFIG,
            output_type=pytesseract.Output.DICT,
        )
        for i, found_text in enumerate(data["text"]):
            if (
                found_text.strip().lower() == text.lower()
                and float(data["conf"][i]) / 100 >= confidence
            ):
                x, y, w, h = (
                    data["left"][i],
                    data["top"][i],
                    data["width"][i],
                    data["height"][i],
                )
                if region:
                    x += region[0]
                    y += region[1]
                return x, y, w, h
        if self.debug:
            self._save_debug_image(img, f"not_found_{text}")
        return None

    def click_text(self, text: str, offset_x: int = 0, offset_y: int = 0) -> bool:
        """Click on found text with optional offset."""
        coords = self.find_text_on_screen(text)
        if not coords:
            logger.error("Text '%s' not found on screen", text)
            if self.debug:
                logger.debug("Saved debug screenshot for '%s'", text)
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
