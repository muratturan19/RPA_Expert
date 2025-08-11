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
    subprocess.run(
        [
            "chrome",
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
            logger.info(
                "Processing date %s with %d transactions", data_entry["tarih"], data_entry["islem_sayisi"]
            )
            # Navigation Phase
            menu_width, _ = pyautogui.size()
            menu_region = (0, 0, menu_width, 120)
            if not self.ocr.click_text(UI_TEXTS["finans_izle"], region=menu_region):
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
