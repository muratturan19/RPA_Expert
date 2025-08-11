"""
Configuration settings for Preston RPA system.
"""

# OCR Settings
OCR_CONFIDENCE = 0.8
# Use Turkish language pack
OCR_LANGUAGE = "tur"
# Tesseract configuration string (single line mode for menu bar)
OCR_TESSERACT_CONFIG = "--psm 7"

# Timing Settings
CLICK_DELAY = 1.0
FORM_FILL_DELAY = 0.5
MODAL_WAIT_TIMEOUT = 10

# Text patterns for OCR
UI_TEXTS = {
    "finans_izle": [
        "Finans - İzle",
        "Finans-İzle",
        "Finans – İzle",
        "Finans — İzle",
        "Finans İzle",
        "FINANS - İZLE",
        "Finans - Izle",
    ],
    "banka_hesap_izleme": "Banka hesap izleme",
    "tamam_button": "Tamam",
    "yeni_button": "Yeni",
    "kaydet_button": "Kaydet",
    "kapat_button": "Kapat",
}

# Mapping tables
BANK_CODES = {
    "233442112": "6293986",  # Account to Bank code mapping
}

CARI_CODES = {
    "default": "120.12.001",  # Default Cari code
}
