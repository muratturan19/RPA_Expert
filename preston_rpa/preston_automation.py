"""Core automation workflow for Preston RPA."""

from __future__ import annotations

import time
from typing import List, Dict

import subprocess
import shutil
from pathlib import Path

import pyautogui
import pygetwindow as gw

from .config import CLICK_DELAY, FORM_FILL_DELAY, UI_TEXTS, BANK_CODES, CARI_CODES
from .logger import get_logger
from .ocr_engine import OCREngine
from .image_matcher import ImageMatcher

logger = get_logger(__name__)


def focus_preston_window(simulator_path: str) -> None:
    """Bring Preston simulator tab to the foreground or open it if missing."""
    target_title = "Preston XI - Kurumsal Kaynak Yönetim Sistemi"

    # Try to locate a window where the desired tab is already active
    preston_windows = gw.getWindowsWithTitle(target_title)
    if preston_windows:
        preston_window = preston_windows[0]
        preston_window.activate()
        preston_window.maximize()
        pyautogui.click(preston_window.center)
        time.sleep(1)
        return

    # If the tab exists but is not active, cycle through Chrome tabs
    chrome_windows = gw.getWindowsWithTitle("Chrome")
    if chrome_windows:
        chrome_window = chrome_windows[0]
        chrome_window.activate()
        chrome_window.maximize()
        for _ in range(10):
            active = gw.getActiveWindow()
            if active and target_title in active.title:
                pyautogui.click(active.center)
                time.sleep(1)
                return
            pyautogui.hotkey("ctrl", "tab")
            time.sleep(0.5)

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
        pyautogui.click(preston_window.center)
    time.sleep(1)


class PrestonRPA:
    def __init__(self):
        self.ocr = OCREngine()
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
            menu_roi = (l + 8, t + 170, r - 8, t + 220)
            center_roi = (l + 200, t + 260, r - 200, t + 420)
            window_rect = (l, t, r - l, b - t)

            found_pair = self.ocr.find_word_pair(window_rect, "finans", "izle")
            found = found_pair or self.ocr.find_text_on_screen(
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
            ImageGrab.grab(bbox=menu_roi).save("debug_menu_roi.png")
            ImageGrab.grab(bbox=center_roi).save("debug_center_roi.png")
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
            window_rect = (window.left, window.top, window.width, window.height)
            if not self.ocr.click_word_pair(window_rect, "Finans", "İzle"):
                raise AssertionError("'Finans - İzle' menu not found")
            time.sleep(CLICK_DELAY)
            if not self.ocr.wait_for_text(UI_TEXTS["banka_hesap_izleme"], timeout=2):
                raise AssertionError("'Finans - İzle' dropdown did not open")
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
