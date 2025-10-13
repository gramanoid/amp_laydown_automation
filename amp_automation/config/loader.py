"""Helpers for loading the consolidated master configuration."""

from __future__ import annotations

import json
from collections.abc import Mapping, MutableMapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any

DEFAULT_CONFIG_FILENAME = "master_config.json"


@dataclass(slots=True)
class Config:
    """Lightweight wrapper around the master configuration payload."""

    data: MutableMapping[str, Any]
    path: Path

    def get(self, key: str, default: Any = None) -> Any:
        return self.data.get(key, default)

    def require(self, key: str) -> Any:
        if key not in self.data:
            raise KeyError(f"Configuration key '{key}' is missing from {self.path}")
        return self.data[key]

    def section(self, key: str) -> Mapping[str, Any]:
        value = self.require(key)
        if not isinstance(value, Mapping):
            raise TypeError(f"Configuration section '{key}' must be a mapping (file: {self.path})")
        return value


def load_master_config(path: str | Path | None = None) -> Config:
    """Load and validate the master configuration file."""

    config_path = (
        Path(path).expanduser().resolve()
        if path
        else Path(__file__).resolve().parents[2] / "config" / DEFAULT_CONFIG_FILENAME
    )

    if not config_path.is_file():
        raise FileNotFoundError(f"Master configuration not found at {config_path}")

    with config_path.open("r", encoding="utf-8-sig") as handle:
        payload = json.load(handle)

    if not isinstance(payload, MutableMapping):
        raise TypeError(f"Configuration root must be a JSON object (file: {config_path})")

    return Config(data=dict(payload), path=config_path)
