"""Logger utility for Preston RPA system."""

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

LOG_FILE = Path(__file__).resolve().parent / "automation.log"
_log_file_cleared = False


def get_logger(name: str = "preston_rpa") -> logging.Logger:
    """Return a configured logger instance.

    The log file is cleared on first logger initialization to ensure a fresh
    log on each application startup.
    """

    global _log_file_cleared
    logger = logging.getLogger(name)

    if not _log_file_cleared and LOG_FILE.exists():
        LOG_FILE.unlink()
        _log_file_cleared = True

    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger
