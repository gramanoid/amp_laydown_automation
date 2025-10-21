"""Aspose Cloud integration helpers for post-processing presentations."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Iterable, Sequence

import requests

logger = logging.getLogger("amp_automation.aspose")

ASPOSE_DEFAULT_BASE_URL = "https://api.aspose.cloud"


class AsposeConfigurationError(RuntimeError):
    """Raised when Aspose credentials are not configured."""


def export_with_aspose(
    pptx_path: str | Path,
    export_formats: Iterable[str],
    output_dir: str | Path,
    *,
    client_id: str | None = None,
    client_secret: str | None = None,
    base_url: str = ASPOSE_DEFAULT_BASE_URL,
    timeout: int = 120,
) -> list[Path]:
    """Convert *pptx_path* into the requested *export_formats* via Aspose Cloud.

    Parameters
    ----------
    pptx_path:
        Source PowerPoint document to convert.
    export_formats:
        Iterable of Aspose-supported target formats (e.g. ``{"pdf", "png"}``).
    output_dir:
        Destination directory for converted assets; created when missing.
    client_id / client_secret:
        OAuth credentials. If omitted, environment variables
        ``ASPOSE_CLIENT_ID`` and ``ASPOSE_CLIENT_SECRET`` are used.
    base_url:
        Aspose Cloud base endpoint. Defaults to the public SaaS host.
    timeout:
        Request timeout in seconds for API calls.

    Returns
    -------
    list[Path]
        Paths to the generated artifacts.
    """

    pptx_path = Path(pptx_path)
    if not pptx_path.is_file():
        raise FileNotFoundError(f"Presentation not found: {pptx_path}")

    export_list = [fmt.lower() for fmt in export_formats]
    if not export_list:
        raise ValueError("export_formats must contain at least one format")

    client_id = client_id or os.getenv("ASPOSE_CLIENT_ID")
    client_secret = client_secret or os.getenv("ASPOSE_CLIENT_SECRET")
    if not client_id or not client_secret:
        raise AsposeConfigurationError(
            "Aspose credentials missing. Provide client_id/client_secret or set "
            "ASPOSE_CLIENT_ID / ASPOSE_CLIENT_SECRET environment variables."
        )

    token = _request_access_token(client_id, client_secret, base_url, timeout)

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    converted_files: list[Path] = []
    with pptx_path.open("rb") as fh:
        file_bytes = fh.read()

    for fmt in export_list:
        target = output_dir / f"{pptx_path.stem}.{fmt}"
        _convert_single_format(
            file_bytes,
            fmt,
            target,
            token,
            base_url,
            timeout,
        )
        converted_files.append(target)

    logger.info(
        "Aspose Cloud generated %d artifact(s) for %s", len(converted_files), pptx_path
    )
    return converted_files


def _request_access_token(
    client_id: str,
    client_secret: str,
    base_url: str,
    timeout: int,
) -> str:
    token_url = f"{base_url.rstrip('/')}/connect/token"
    response = requests.post(
        token_url,
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
        },
        timeout=timeout,
    )
    if response.status_code != 200:
        raise RuntimeError(
            f"Aspose token request failed ({response.status_code}): {response.text}"
        )

    payload = response.json()
    token = payload.get("access_token")
    if not token:
        raise RuntimeError("Aspose token response missing access_token")
    return token


def _convert_single_format(
    file_bytes: bytes,
    fmt: str,
    target_path: Path,
    token: str,
    base_url: str,
    timeout: int,
) -> None:
    url = f"{base_url.rstrip('/')}/v3.0/slides/convert/{fmt}"
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.post(
        url,
        headers=headers,
        files={"document": (target_path.name, file_bytes)},
        timeout=timeout,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Aspose conversion to '{fmt}' failed ({response.status_code}): {response.text}"
        )

    target_path.write_bytes(response.content)
