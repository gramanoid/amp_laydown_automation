"""Template V4 geometry constants shared across presentation modules."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict

from pptx.util import Inches


TEMPLATE_V4_COLUMN_WIDTHS_EMU: list[int] = [
    1_000_000,  # Campaign column - widened to ~1.09 inches to fit words like "CONDITION"
    729_251,
    831_384,
    338_274,
    400_567,
    400_567,
    400_567,
    414_770,
    415_954,
    465_506,
    437_595,
    443_865,
    400_567,
    437_595,
    352_043,
    449_092,
    400_567,
    400_567,
]

TEMPLATE_V4_COLUMN_WIDTHS_INCHES: list[float] = [
    width / 914_400 for width in TEMPLATE_V4_COLUMN_WIDTHS_EMU
]

TEMPLATE_V4_ROW_HEIGHT_HEADER_EMU = 161_729
TEMPLATE_V4_ROW_HEIGHT_BODY_EMU = 99_205
TEMPLATE_V4_ROW_HEIGHT_TRAILER_EMU = 0

TEMPLATE_V4_ROW_HEIGHT_HEADER_INCHES = TEMPLATE_V4_ROW_HEIGHT_HEADER_EMU / 914_400
TEMPLATE_V4_ROW_HEIGHT_BODY_INCHES = TEMPLATE_V4_ROW_HEIGHT_BODY_EMU / 914_400


@dataclass(frozen=True)
class TemplateTableBounds:
    """Store Template V4 table position and size in inches."""

    left: float
    top: float
    width: float
    height: float

    def as_dict(self) -> Dict[str, float]:
        return {
            "left": self.left,
            "top": self.top,
            "width": self.width,
            "height": self.height,
        }

    def as_inches(self) -> Dict[str, Inches]:
        return {
            "left": Inches(self.left),
            "top": Inches(self.top),
            "width": Inches(self.width),
            "height": Inches(self.height),
        }


TEMPLATE_V4_TABLE_LEFT_EMU = 163_582
TEMPLATE_V4_TABLE_TOP_EMU = 638_117
TEMPLATE_V4_TABLE_WIDTH_EMU = 8_531_095
TEMPLATE_V4_TABLE_HEIGHT_EMU = 3_766_424

TEMPLATE_V4_TABLE_BOUNDS = TemplateTableBounds(
    left=TEMPLATE_V4_TABLE_LEFT_EMU / 914_400,
    top=TEMPLATE_V4_TABLE_TOP_EMU / 914_400,
    width=TEMPLATE_V4_TABLE_WIDTH_EMU / 914_400,
    height=TEMPLATE_V4_TABLE_HEIGHT_EMU / 914_400,
)


__all__ = [
    "TEMPLATE_V4_COLUMN_WIDTHS_EMU",
    "TEMPLATE_V4_COLUMN_WIDTHS_INCHES",
    "TEMPLATE_V4_ROW_HEIGHT_HEADER_EMU",
    "TEMPLATE_V4_ROW_HEIGHT_BODY_EMU",
    "TEMPLATE_V4_ROW_HEIGHT_TRAILER_EMU",
    "TEMPLATE_V4_ROW_HEIGHT_HEADER_INCHES",
    "TEMPLATE_V4_ROW_HEIGHT_BODY_INCHES",
    "TEMPLATE_V4_TABLE_LEFT_EMU",
    "TEMPLATE_V4_TABLE_TOP_EMU",
    "TEMPLATE_V4_TABLE_WIDTH_EMU",
    "TEMPLATE_V4_TABLE_HEIGHT_EMU",
    "TEMPLATE_V4_TABLE_BOUNDS",
    "TemplateTableBounds",
]
