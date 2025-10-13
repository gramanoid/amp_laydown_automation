"""Shared logging helpers."""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Optional


def configure_logger(log_directory: Path, prefix: str = "automation") -> logging.Logger:
    """Configure a module-level logger writing to the provided directory."""

    log_directory.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = log_directory / f"{prefix}_{timestamp}.log"

    logger = logging.getLogger("amp_automation")
    logger.setLevel(logging.DEBUG)

    if logger.hasHandlers():
        logger.handlers.clear()

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s - %(levelname)s - %(name)s - %(message)s")
    )
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(
        logging.Formatter("%(levelname)s - %(message)s")
    )
    logger.addHandler(console_handler)

    logger.info("Logging initialised")
    logger.debug("Log file: %s", log_path)

    return logger
