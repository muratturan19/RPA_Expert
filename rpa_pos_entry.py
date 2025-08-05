import argparse
import logging
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any

from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("rpa.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def read_excel(path: Path) -> List[Dict[str, Any]]:
    """Read POS data from an Excel file."""
    logger.info("Reading Excel data from %s", path)
    workbook = load_workbook(path)
    sheet = workbook.active
    rows: List[Dict[str, Any]] = []

    column_map = {
        "Tarih": "tarih",
        "Firma": "firma",
        "Tutar": "tutar",
        "Açıklama": "aciklama",
        "Döviz": "doviz",
        "Vade Tarihi": "vade_tarihi",
    }

    raw_header = [cell for cell in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
    header = [column_map.get(cell, cell) for cell in raw_header]

    for row in sheet.iter_rows(min_row=2, values_only=True):
        item: Dict[str, Any] = {}
        for key, value in zip(header, row):
            if isinstance(value, datetime):
                item[key] = value.date().isoformat()
            else:
                item[key] = value
        rows.append(item)
    logger.info("Loaded %d rows from Excel", len(rows))
    return rows


def setup_driver() -> webdriver.Chrome:
    """Initialise Chrome WebDriver."""
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    return driver


def open_application(driver: webdriver.Chrome, html_path: Path) -> None:
    """Open local Preston simulator."""
    driver.get(html_path.resolve().as_uri())
    logger.info("Opened Preston simulator at %s", html_path)


def navigate_to_pos(driver: webdriver.Chrome) -> None:
    """Navigate to Finance > Tahsilat and open POS entry modal."""
    menu_selector = "div.menu-item[data-menu='finans-tahsilat']"
    pos_icon_selector = "div.ribbon-icon[data-tooltip='Pos Giri\u015fi']"
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, menu_selector))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, pos_icon_selector))).click()
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "posModal")))
    logger.debug("POS modal opened")


def fill_pos_form(driver: webdriver.Chrome, data: Dict[str, Any]) -> None:
    """Fill POS entry form and save."""
    driver.find_element(By.ID, "pos-tarih").clear()
    driver.find_element(By.ID, "pos-tarih").send_keys(str(data.get("tarih", "")))

    driver.find_element(By.ID, "pos-firma").send_keys(str(data.get("firma", "")))
    driver.find_element(By.ID, "pos-aciklama").send_keys(str(data.get("aciklama", "")))
    driver.find_element(By.ID, "pos-tutar").send_keys(str(data.get("tutar", "")))

    doviz = data.get("doviz")
    if doviz:
        driver.find_element(By.ID, "posDoviz").send_keys(str(doviz))

    vade = data.get("vade_tarihi")
    if vade:
        driver.find_element(By.ID, "posVadeTarihi").send_keys(str(vade))

    driver.find_element(By.CSS_SELECTOR, ".modal-buttons .primary").click()
    WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.ID, "modalOverlay")))
    logger.info("POS entry saved for %s", data)


def process_entries(driver: webdriver.Chrome, entries: List[Dict[str, Any]]) -> None:
    for entry in entries:
        try:
            navigate_to_pos(driver)
            fill_pos_form(driver, entry)
        except Exception as exc:  # pylint: disable=broad-except
            logger.exception("Error processing row %s: %s", entry, exc)


def main(excel_path: Path) -> None:
    data_rows = read_excel(excel_path)
    driver = setup_driver()
    try:
        html_file = Path(__file__).parent / "RPA_Expert.html"
        open_application(driver, html_file)
        process_entries(driver, data_rows)
    except WebDriverException as exc:
        logger.exception("WebDriver error: %s", exc)
    finally:
        driver.quit()
        logger.info("Driver closed")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Automate POS entries on Preston simulator")
    parser.add_argument("--excel", default="pos_data.xlsx", help="Path to Excel file with POS data")
    args = parser.parse_args()
    main(Path(args.excel))
