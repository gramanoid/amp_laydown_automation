"""Media-related utility helpers."""

from __future__ import annotations

__all__ = ["normalize_media_type"]


def normalize_media_type(media_type: object) -> str:
    """Standardize raw media type values for consistent comparisons."""

    if media_type is None:
        return ""

    media_type_str = str(media_type).strip().upper()
    if media_type_str in {"TV", "TELEVISION"}:
        return "Television"
    if media_type_str == "DIGITAL":
        return "Digital"
    return media_type_str
