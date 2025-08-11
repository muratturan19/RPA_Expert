"""OCR engine using pytesseract for screen text detection."""

from __future__ import annotations

import time
import re
from datetime import datetime
from typing import Optional, Tuple, Iterable
from pathlib import Path

import numpy as np
import pytesseract
import pyautogui
import pygetwindow as gw
import cv2
from PIL import Image

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
                time.sleep(0.3)
            img = pyautogui.screenshot(region=region)
            img = self._preprocess_image(img)
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

    @staticmethod
    def _preprocess_image(img: Image.Image) -> Image.Image:
        gray = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)
        gray = cv2.resize(gray, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_LINEAR)
        thresh = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2
        )
        return Image.fromarray(thresh)

    @staticmethod
    def _normalize(text: str) -> str:
        mapping = str.maketrans({
            "İ": "I",
            "ı": "i",
            "Ş": "S",
            "ş": "s",
            "Ğ": "G",
            "ğ": "g",
            "Ç": "C",
            "ç": "c",
            "Ö": "O",
            "ö": "o",
            "Ü": "U",
            "ü": "u",
        })
        text = text.translate(mapping)
        text = re.sub(r"[-–—]", "-", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip().casefold()

    def find_text_on_screen(
        self,
        text: Iterable[str] | str,
        region=None,
        confidence: float = OCR_CONFIDENCE,
        normalize: bool = True,
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
        variants = [text] if isinstance(text, str) else list(text)
        targets = [
            self._normalize(v) if normalize else v.casefold()
            for v in variants
        ]
        lines = {}
        for i, found_text in enumerate(data["text"]):
            if not found_text.strip():
                continue
            if float(data["conf"][i]) / 100 < confidence:
                continue
            key = (
                data["page_num"][i],
                data["block_num"][i],
                data["par_num"][i],
                data["line_num"][i],
            )
            line = lines.setdefault(
                key,
                {"words": [], "left": [], "top": [], "right": [], "bottom": []},
            )
            line["words"].append(found_text)
            left, top, width, height = (
                data["left"][i],
                data["top"][i],
                data["width"][i],
                data["height"][i],
            )
            line["left"].append(left)
            line["top"].append(top)
            line["right"].append(left + width)
            line["bottom"].append(top + height)
        for line in lines.values():
            line_text = " ".join(line["words"])
            line_norm = self._normalize(line_text) if normalize else line_text.casefold()
            if line_norm in targets:
                x = min(line["left"])
                y = min(line["top"])
                w = max(line["right"]) - x
                h = max(line["bottom"]) - y
                if region:
                    x += region[0]
                    y += region[1]
                return x, y, w, h
        if self.debug and variants:
            miss = self._normalize(variants[0]) if normalize else variants[0].casefold()
            self._save_debug_image(img, f"not_found_{miss}")
        return None

    def click_text(
        self,
        text: Iterable[str] | str,
        offset_x: int = 0,
        offset_y: int = 0,
        region=None,
    ) -> bool:
        """Click on found text or any of its variants."""
        variants = [text] if isinstance(text, str) else list(text)
        for variant in variants:
            coords = self.find_text_on_screen(variant, region=region)
            if coords:
                x, y, w, h = coords
                pyautogui.click(x + w // 2 + offset_x, y + h // 2 + offset_y)
                time.sleep(0.1)
                return True
        logger.error("Text '%s' not found on screen", text)
        if self.debug:
            logger.debug("Saved debug screenshot for '%s'", text)
        return False

    def wait_for_text(
        self, text: str, timeout: float = 10, region=None
    ) -> bool:
        """Wait until text appears on screen."""
        end_time = time.time() + timeout
        while time.time() < end_time:
            if self.find_text_on_screen(text, region=region):
                return True
            time.sleep(0.5)
        logger.error("Timeout waiting for text: %s", text)
        return False
