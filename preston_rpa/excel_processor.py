"""Excel processing module for Preston RPA."""

from __future__ import annotations

import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List

from openpyxl import load_workbook

from .logger import get_logger

logger = get_logger(__name__)

POSH_PATTERN = re.compile(r"POSH.*\d{5,}$")


def _parse_date(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, str):
        try:
            return datetime.strptime(value, "%d.%m.%Y").strftime("%d.%m.%Y")
        except ValueError:
            try:
                return datetime.strptime(value, "%Y-%m-%d").strftime("%d.%m.%Y")
            except ValueError:
                logger.warning("Unknown date format: %s", value)
    return ""


def process_excel_file(file_path: Path | str) -> List[Dict[str, object]]:
    """Extract and process transaction data from Excel.

    Parameters
    ----------
    file_path: Path or str
        Path to the Excel file.

    Returns
    -------
    List[Dict[str, object]]
        Structured data for RPA processing grouped by date.
    """
    file_path = Path(file_path)
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    account_no = ws["B6"].value
    if not account_no:
        raise ValueError("Account number (B6) not found in Excel file")

    groups: Dict[str, Dict[str, object]] = defaultdict(lambda: {"toplam_tutar": 0, "islem_sayisi": 0})

    for row in ws.iter_rows(min_row=23, values_only=True):
        if len(row) < 5:
            continue
        islem_tarihi, _, aciklama, islem_tutar, _ = row[:5]
        if not aciklama or not POSH_PATTERN.search(str(aciklama)):
            continue
        tarih = _parse_date(islem_tarihi)
        if not tarih:
            continue
        try:
            amount = float(islem_tutar)
        except (TypeError, ValueError):
            logger.warning("Invalid amount %s on %s", islem_tutar, tarih)
            continue
        data = groups[tarih]
        data["toplam_tutar"] += amount
        data["islem_sayisi"] += 1

    results = [
        {
            "hesap_no": str(account_no),
            "tarih": tarih,
            "toplam_tutar": round(info["toplam_tutar"], 2),
            "islem_sayisi": info["islem_sayisi"],
            "aciklama": "GÜNLÜK POS TAHSİLATI",
        }
        for tarih, info in sorted(groups.items())
    ]
    logger.info("Processed %d date groups from %s", len(results), file_path)
    return results
