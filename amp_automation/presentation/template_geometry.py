"""Template V4 geometry constants shared across presentation modules."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict

from pptx.util import Inches


TEMPLATE_V4_COLUMN_WIDTHS_INCHES: list[float] = [
    0.888,
    0.798,
    0.909,
    0.370,
    0.438,
    0.438,
    0.438,
    0.454,
    0.455,
    0.509,
    0.479,
    0.485,
    0.438,
    0.479,
    0.385,
    0.491,
    0.438,
    0.438,
]


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


TEMPLATE_V4_TABLE_BOUNDS = TemplateTableBounds(
    left=0.179,
    top=0.698,
    width=9.33,
    height=4.119,
)


__all__ = [
    "TEMPLATE_V4_COLUMN_WIDTHS_INCHES",
    "TEMPLATE_V4_TABLE_BOUNDS",
    "TemplateTableBounds",
]
