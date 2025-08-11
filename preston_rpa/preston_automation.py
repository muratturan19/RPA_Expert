"""Core automation workflow for Preston RPA."""

from __future__ import annotations

import time
from typing import List, Dict

import subprocess
import pyautogui
import pygetwindow as gw

from .config import CLICK_DELAY, FORM_FILL_DELAY, UI_TEXTS, BANK_CODES, CARI_CODES
from .logger import get_logger
from .ocr_engine import OCREngine
from .image_matcher import ImageMatcher

logger = get_logger(__name__)


def focus_preston_window(simulator_path: str) -> None:
    """Bring Preston window to foreground or open it if missing."""
    preston_windows = gw.getWindowsWithTitle("Preston")
    if preston_windows:
        preston_window = preston_windows[0]
        preston_window.activate()
        preston_window.maximize()
        pyautogui.click(preston_window.center)
    else:
        subprocess.run(
            [
                "chrome",
                "--new-window",
                "--start-maximized",
                f"file:///{simulator_path}",
            ]
        )
        time.sleep(2)
        preston_windows = gw.getWindowsWithTitle("Preston")
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
        time.sleep(2)
        for entry in excel_data:
            if not self.running:
                break
            self.execute_workflow(entry)
        logger.info("Automation finished")

    def stop(self):
        self.running = False

    # The following methods are placeholders demonstrating the sequence.
    def execute_workflow(self, data_entry: Dict[str, object]):
        """Execute simplified Preston workflow for a single date group."""
        try:
            logger.info("Processing date %s with %d transactions", data_entry["tarih"], data_entry["islem_sayisi"])
            # Navigation Phase
            self.ocr.click_text(UI_TEXTS["finans_menu"])
            time.sleep(CLICK_DELAY)
            self.ocr.click_text(UI_TEXTS["izle_tab"])
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
