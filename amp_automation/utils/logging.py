"""Shared logging helpers."""

from __future__ import annotations

import logging
import os
from datetime import datetime
from pathlib import Path


def configure_logger(
    log_directory: Path,
    prefix: str = "automation",
    *,
    default_level: str = "INFO",
    console_enabled: bool = True,
    file_enabled: bool = True,
    console_level_floor: int | None = logging.INFO,
) -> logging.Logger:
    """Configure or disable the project logger based on config/environment.

    ``default_level`` is the fallback log level when no explicit override is present.
    ``console_enabled``/``file_enabled`` toggle their respective handlers. The
    ``console_level_floor`` parameter keeps console output at or above a minimum
    severity (INFO by default) while still honouring explicit overrides such as
    ``AMP_AUTOMATION_LOG_LEVEL=DEBUG``.
    """

    logger = logging.getLogger("amp_automation")
    logger.handlers.clear()
    logger.propagate = False

    disable_tokens = {"OFF", "NONE", "DISABLED"}

    env_raw = os.getenv("AMP_AUTOMATION_LOG_LEVEL")
    env_override_name: str | None = None
    if env_raw is not None:
        stripped = env_raw.strip()
        if stripped:
            env_override_name = stripped.upper()
        else:
            env_override_name = None

    default_level_name = (default_level or "OFF").strip().upper()
    if not default_level_name:
        default_level_name = "OFF"

    if env_override_name and env_override_name in disable_tokens:
        logger.disabled = True
        return logger

    if env_override_name is None and default_level_name in disable_tokens:
        logger.disabled = True
        return logger

    logger.disabled = False

    level_name = env_override_name or default_level_name or "INFO"
    try:
        level = getattr(logging, level_name)
    except AttributeError:
        level = logging.INFO

    logger.setLevel(level)

    log_directory.mkdir(parents=True, exist_ok=True)
    # Explicitly use local system time (not UTC)
    timestamp = datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")
    log_path = log_directory / f"{prefix}_{timestamp}.log"

    if file_enabled:
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        file_handler.setLevel(level)
        file_handler.setFormatter(
            logging.Formatter("%(asctime)s - %(levelname)s - %(name)s - %(message)s")
        )
        logger.addHandler(file_handler)

    if console_enabled:
        console_handler = logging.StreamHandler()
        console_level = level
        if (
            env_override_name is None
            and console_level_floor is not None
            and console_level < console_level_floor
        ):
            console_level = console_level_floor
        console_handler.setLevel(console_level)
        console_handler.setFormatter(logging.Formatter("%(levelname)s - %(message)s"))
        logger.addHandler(console_handler)

    logger.info("Logging initialised")
    if file_enabled:
        logger.debug("Log file: %s", log_path)

    return logger
