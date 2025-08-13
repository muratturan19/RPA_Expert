"""Core automation workflow for Preston RPA."""

from __future__ import annotations

import time
from typing import List, Dict

import subprocess
import shutil
from pathlib import Path

import pyautogui
import pygetwindow as gw

from .config import (
    CLICK_DELAY,
    FORM_FILL_DELAY,
    UI_TEXTS,
    BANK_CODES,
    CARI_CODES,
    OCR_CONFIDENCE,
)
from .logger import get_logger
from .ocr_engine import OCREngine
from .image_matcher import ImageMatcher

logger = get_logger(__name__)


def focus_preston_window(simulator_path: str) -> None:
    """Bring Preston simulator tab to the foreground or open it if missing."""
    target_title = "Preston Xi Kurumsal Kay"

    # Try to locate a window where the desired tab is already active
    preston_windows = gw.getWindowsWithTitle(target_title)
    if preston_windows:
        preston_window = preston_windows[0]
        preston_window.activate()
        preston_window.maximize()
        time.sleep(1)
        return

    # Open the simulator if the tab could not be found
    chrome_paths = [
        shutil.which("chrome"),
        shutil.which("google-chrome"),
        shutil.which("chromium"),
        r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
    ]

    chrome_executable = next((p for p in chrome_paths if p and Path(p).exists()), None)
    if not chrome_executable:
        raise FileNotFoundError("Chrome executable not found. Please install Google Chrome or add it to PATH.")

    subprocess.run(
        [
            chrome_executable,
            "--new-window",
            "--start-maximized",
            f"file:///{simulator_path}",
        ]
    )
    time.sleep(2)
    preston_windows = gw.getWindowsWithTitle(target_title)
    if preston_windows:
        preston_window = preston_windows[0]
        preston_window.activate()
        preston_window.maximize()
    time.sleep(1)


class PrestonRPA:
    def __init__(self):
        self.ocr = OCREngine(debug=True)
        self.image_matcher = ImageMatcher()
        self.running = True

    def start_automation(self, excel_data: List[Dict[str, object]], simulator_path: str):
        """Main automation workflow."""
        logger.info("Starting automation for %d date groups", len(excel_data))
        focus_preston_window(simulator_path)
        if not self._wait_for_preston_ready():
            logger.error("Preston not ready")
            return
        for entry in excel_data:
            if not self.running:
                break
            self.execute_workflow(entry)
        logger.info("Automation finished")

    def stop(self):
        self.running = False

    def _wait_for_preston_ready(self, timeout: float = 15) -> bool:
        """Preston sekmesi görünür/aktif ve içerik çizilmiş olmadan True dönmez."""
        import uiautomation as auto
        from PIL import ImageGrab

        CENTER = ["Preston Banka Hesap İzleme", "Banka Hesap İzleme"]

        end_time = time.time() + timeout
        ok_streak = 0

        def _chrome():
            return auto.WindowControl(searchDepth=1, NameRe=r".*Chrome")

        while time.time() < end_time:
            ch = _chrome()
            if not ch.Exists(0.3):
                time.sleep(0.3)
                continue

            ch.SetActive()
            tab = ch.TabItemControl(NameRe=r"(Preston\s+X[iI]\b.*Kurumsal.*|Preston\.html)")
            if tab.Exists(0.2):
                try:
                    tab.Select()
                except Exception:
                    pass

            l, t, r, b = ch.BoundingRectangle
            menu_left, menu_top = l + 8, t + 170
            menu_width, menu_height = (r - 8) - menu_left, (t + 220) - menu_top
            menu_roi = (menu_left, menu_top, menu_width, menu_height)
            center_left, center_top = l + 200, t + 260
            center_width, center_height = (r - 200) - center_left, (t + 420) - center_top
            center_roi = (center_left, center_top, center_width, center_height)
            window_rect = (l, t, r - l, b - t)

            found_izle = self.ocr.find_text_on_screen(
                ["İzle", "izle", "IZLE"], region=menu_roi
            )
            found = found_izle or self.ocr.find_text_on_screen(
                CENTER, region=center_roi, normalize=True
            )

            if found:
                ok_streak += 1
                if ok_streak >= 2:
                    return True
            else:
                ok_streak = 0

            time.sleep(0.25)

        try:
            ImageGrab.grab(
                bbox=(
                    menu_left,
                    menu_top,
                    menu_left + menu_width,
                    menu_top + menu_height,
                )
            ).save("debug_menu_roi.png")
            ImageGrab.grab(
                bbox=(
                    center_left,
                    center_top,
                    center_left + center_width,
                    center_top + center_height,
                )
            ).save("debug_center_roi.png")
        except Exception:
            pass
        logger.error("Preston ready check failed; ROI screenshots saved.")
        return False

    # The following methods are placeholders demonstrating the sequence.
    def execute_workflow(self, data_entry: Dict[str, object]):
        """Execute simplified Preston workflow for a single date group."""
        try:
            logger.info(
                "Processing date %s with %d transactions", data_entry["tarih"], data_entry["islem_sayisi"]
            )
            # Navigation Phase
            window = gw.getActiveWindow()
            if not window:
                raise AssertionError("Preston window not active")
            logger.info(
                "Window: left=%d, top=%d, width=%d, height=%d",
                window.left,
                window.top,
                window.width,
                window.height,
            )
            screen_w, screen_h = pyautogui.size()
            window_rect = (0, 0, window.width, window.height)
            if (
                window.left < 0
                or window.top < 0
                or window.left + window.width > screen_w
                or window.top + window.height > screen_h
            ):
                logger.warning(
                    "Window rect out of bounds; using screen bounds instead"
                )
                window_rect = (0, 0, min(window.width, screen_w), min(window.height, screen_h))
            logger.info("Window rect: %s", window_rect)
            # Precise menu region covering the "Finans - İzle" menu
            menu_region = (300, 100, 500, 200)
            logger.info(f"Menu region: {menu_region}")
            menu_screenshot = pyautogui.screenshot(region=menu_region)
            menu_screenshot.save("debug_menu_only.png")
            self.ocr._save_debug_image(menu_screenshot, "debug_menu_region")
            # Menu search screenshots
            self.ocr._screenshot(region=menu_region, step_name="menu_search_before")
            deadline = time.time() + 3
            ok = 0
            while time.time() < deadline:
                bbox = self.ocr.find_text_on_screen(
                    ["İzle", "izle", "Izle"],
                    region=menu_region,
                    confidence=0.6,
                )
                if bbox:
                    ok += 1
                    if ok >= 2:
                        break
                else:
                    ok = 0
                time.sleep(0.2)
            else:
                raise AssertionError("'İzle' görünmedi")
            bbox = self.ocr.find_text_on_screen(
                ["İzle", "izle", "IZLE"],
                region=menu_region,
                confidence=OCR_CONFIDENCE,
            )
            if bbox:
                x, y, w, h = bbox
                pyautogui.click(x + w // 2, y + h // 2)
                logger.info("Successfully clicked İzle menu")
            else:
                raise AssertionError("'İzle' menu not found")
            self.ocr._screenshot(region=menu_region, step_name="menu_search_after")
            time.sleep(CLICK_DELAY)
            if not self.ocr.wait_for_text(
                ["Banka Hesap İzleme", "Banka hesap izleme"],
                timeout=2,
                confidence=0.6,
            ):
                raise AssertionError("'Finans - İzle' dropdown did not open")
            self.ocr._screenshot(region=window_rect, step_name="menu_after_dropdown")
            time.sleep(CLICK_DELAY)
            # Placeholder for more navigation steps...

            # Account Selection
            # In real application, we would type account number and confirm dialogs
            logger.debug("Selecting account %s", data_entry["hesap_no"])

            # Date Setup
            logger.debug("Setting date range %s", data_entry["tarih"])

            # Form Creation and Data Entry
            logger.debug("Entering amount %.2f", data_entry["toplam_tutar"])
            time.sleep(FORM_FILL_DELAY)
            logger.debug("Entering description %s", data_entry["aciklama"])
            time.sleep(FORM_FILL_DELAY)

            logger.info("Entry for %s completed", data_entry["tarih"])
        except Exception as exc:
            logger.exception("Workflow failed for %s: %s", data_entry.get("tarih"), exc)
