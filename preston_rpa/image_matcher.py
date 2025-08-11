"""Image recognition utilities using OpenCV."""

from __future__ import annotations

from typing import Optional, Tuple

import cv2
import numpy as np
import pyautogui
from PIL import ImageGrab

from .logger import get_logger

logger = get_logger(__name__)


class ImageMatcher:
    def _screenshot(self):
        try:
            img = ImageGrab.grab()
            return cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
        except Exception as exc:
            logger.error("Screenshot failed: %s", exc)
            return None

    def find_icon(self, template_path: str, confidence: float = 0.9) -> Optional[Tuple[int, int, int, int]]:
        """Find UI icons using template matching."""
        screen = self._screenshot()
        if screen is None:
            return None
        template = cv2.imread(template_path, 0)
        if template is None:
            logger.error("Template not found: %s", template_path)
            return None
        res = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        if max_val < confidence:
            return None
        h, w = template.shape
        return max_loc[0], max_loc[1], w, h

    def click_icon(self, template_path: str, confidence: float = 0.9) -> bool:
        coords = self.find_icon(template_path, confidence)
        if not coords:
            logger.error("Icon not found: %s", template_path)
            return False
        x, y, w, h = coords
        pyautogui.click(x + w // 2, y + h // 2)
        return True
