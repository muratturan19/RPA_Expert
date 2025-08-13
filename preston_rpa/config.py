"""
Configuration settings for Preston RPA system.
"""

# OCR Settings
# Accept slightly lower confidence OCR matches to improve robustness
OCR_CONFIDENCE = 0.6
# Use Turkish language pack
OCR_LANGUAGE = "tur"
# Tesseract configuration string optimized for single-line menu text
OCR_TESSERACT_CONFIG = (
    "--oem 1 --psm 7 "
    "-c preserve_interword_spaces=1 "
    "-c tessedit_char_whitelist=ABCÇDEFGĞHIİJKLMNOÖPQRSŞTUÜVWXYZabcçdefgğhıijklmnoöpqrsştuüvwxyz0123456789 |-"
)
# Minimum similarity ratio (0-1) for fuzzy text matching in OCR
OCR_FUZZY_THRESHOLD = 0.65

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
    "banka_hesap_izleme": ["Banka hesap izleme", "banka hesap izleme"],
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
