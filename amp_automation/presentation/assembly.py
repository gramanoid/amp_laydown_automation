"""
Excel to PowerPoint Automation Script - GEOGRAPHY VERSION
This version uses the 'Plan - Geography' column instead of '**Country' column
to extract market/country information from the hierarchical geography string.

Changes from FINAL_PRODUCTION.py:
- Uses 'Plan - Geography' column (column K) instead of '**Country'
- Same extraction logic: splits by " | " and takes the last part
- Example: "Global | EMEA | MEA | Pakistan" → "Pakistan"
"""
import os
import logging
import traceback
from datetime import datetime
import pandas as pd
import numpy as np
import ast, pathlib, inspect, textwrap
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL, MSO_FILL_TYPE

from amp_automation.config import Config, load_master_config
from amp_automation.data import (
    get_month_specific_tv_metrics,
    load_and_prepare_data as modular_load_and_prepare_data,
)
from amp_automation.presentation.charts import (
    ChartStyleContext,
    add_pie_chart as presentation_add_pie_chart,
    prepare_campaign_type_chart_data,
    prepare_funnel_chart_data,
    prepare_media_type_chart_data,
)
from amp_automation.presentation.tables import (
    CellStyleContext,
    TableLayout,
    add_and_style_table as presentation_add_and_style_table,
    ensure_font_consistency as _ensure_font_consistency,
)
from amp_automation.utils.media import normalize_media_type

def _rgb_color(config_value, fallback):
    try:
        r, g, b = config_value
        return RGBColor(int(r), int(g), int(b))
    except (TypeError, ValueError):  # pragma: no cover - defensive fallback
        return RGBColor(*fallback)


def _coord_from_config(config_mapping, fallback):
    return {
        "left": float(config_mapping.get("left_inches", config_mapping.get("left", fallback["left"]))),
        "top": float(config_mapping.get("top_inches", config_mapping.get("top", fallback["top"]))),
        "width": float(config_mapping.get("width_inches", config_mapping.get("width", fallback["width"]))),
        "height": float(config_mapping.get("height_inches", config_mapping.get("height", fallback["height"]))),
}


MARGIN_EMU_LR = 45720  # Approx Pt(3.6), from template analysis for left/right cell margins
ZERO_THRESHOLD = 0.01  # Values below this (absolute) are treated as zero for display/coloring


def _initialize_from_config(config: Config) -> None:
    global MASTER_CONFIG
    global presentation_config, fonts_config, font_sizes
    global colors_config, media_colors_config, ui_colors_config
    global table_config, row_heights_config, table_position_config
    global comments_config, comments_title_pos_config, comments_box_pos_config
    global title_position_config, charts_config, chart_positions_config
    global TABLE_PLACEHOLDER_NAME
    global CLR_BLACK, CLR_WHITE, CLR_LIGHT_GRAY_TEXT, CLR_TABLE_GRAY, CLR_HEADER_GREEN
    global CLR_COMMENTS_GRAY, CLR_SUBTOTAL_GRAY
    global CLR_TELEVISION, CLR_DIGITAL, CLR_OOH, CLR_OTHER
    global DEFAULT_FONT_NAME, FONT_SIZE_HEADER, FONT_SIZE_BODY
    global FONT_SIZE_CHART_TITLE, FONT_SIZE_CHART_LABELS
    global FONT_SIZE_TITLE, FONT_SIZE_LEGEND, FONT_SIZE_COMMENTS
    global TABLE_ROW_HEIGHT_HEADER, TABLE_ROW_HEIGHT_BODY, TABLE_ROW_HEIGHT_SUBTOTAL
    global TABLE_COLUMN_WIDTHS, TABLE_TOP_OVERRIDE
    global TABLE_CELL_STYLE_CONTEXT
    global CHART_STYLE_CONTEXT, CHART_COLOR_MAPPING, CHART_COLOR_CYCLE
    global MAX_ROWS_PER_SLIDE, SPLIT_STRATEGY, SHOW_CHARTS_ON_SPLITS, SHOW_CARRIED_SUBTOTAL, CONTINUATION_INDICATOR
    global ELEMENT_COORDINATES

    MASTER_CONFIG = config

    presentation_config = MASTER_CONFIG.section("presentation")
    fonts_config = presentation_config.get("fonts", {})
    font_sizes = fonts_config.get("sizes_pt", {})
    colors_config = presentation_config.get("colors", {})
    media_colors_config = colors_config.get("media_types", {})
    ui_colors_config = colors_config.get("ui_elements", {})
    table_config = presentation_config.get("table", {})
    row_heights_config = table_config.get("row_heights", {})
    table_position_config = table_config.get("positioning", {})
    comments_config = presentation_config.get("comments", {})
    comments_title_pos_config = comments_config.get("title_positioning", {})
    comments_box_pos_config = comments_config.get("box_positioning", {})
    title_position_config = presentation_config.get("title", {}).get("positioning", {})
    charts_config = presentation_config.get("charts", {})
    chart_positions_config = charts_config.get("positioning", {})

    TABLE_PLACEHOLDER_NAME = table_config.get("placeholder_name", "Table Placeholder 1")

    CLR_BLACK = RGBColor(0, 0, 0)
    CLR_WHITE = RGBColor(255, 255, 255)
    CLR_LIGHT_GRAY_TEXT = _rgb_color(ui_colors_config.get("light_gray_text", {}).get("rgb"), (191, 191, 191))
    CLR_TABLE_GRAY = _rgb_color(ui_colors_config.get("table_gray", {}).get("rgb"), (191, 191, 191))
    CLR_HEADER_GREEN = _rgb_color(ui_colors_config.get("header_green", {}).get("rgb"), (56, 236, 4))
    CLR_COMMENTS_GRAY = _rgb_color(ui_colors_config.get("comments_gray", {}).get("rgb"), (242, 242, 242))
    CLR_SUBTOTAL_GRAY = _rgb_color(ui_colors_config.get("subtotal_gray", {}).get("rgb"), (217, 217, 217))

    CLR_TELEVISION = _rgb_color(media_colors_config.get("television", {}).get("rgb"), (113, 212, 141))
    CLR_DIGITAL = _rgb_color(media_colors_config.get("digital", {}).get("rgb"), (253, 242, 183))
    CLR_OOH = _rgb_color(media_colors_config.get("ooh", {}).get("rgb"), (255, 191, 0))
    CLR_OTHER = _rgb_color(media_colors_config.get("other", {}).get("rgb"), (176, 211, 255))

    DEFAULT_FONT_NAME = fonts_config.get("default_family", "Calibri")
    FONT_SIZE_HEADER = Pt(float(font_sizes.get("header", 7.5)))
    FONT_SIZE_BODY = Pt(float(font_sizes.get("body", 7.0)))
    FONT_SIZE_CHART_TITLE = Pt(float(font_sizes.get("chart_title", 8.0)))
    FONT_SIZE_CHART_LABELS = Pt(float(font_sizes.get("chart_labels", 6.0)))
    FONT_SIZE_TITLE = Pt(float(font_sizes.get("title", 11.0)))
    FONT_SIZE_LEGEND = Pt(float(font_sizes.get("legend", 6.0)))
    FONT_SIZE_COMMENTS = Pt(float(font_sizes.get("comments", 9.0)))

    TABLE_ROW_HEIGHT_HEADER = Pt(float(row_heights_config.get("header_inches", 0.139)) * 72)
    TABLE_ROW_HEIGHT_BODY = Pt(float(row_heights_config.get("body_inches", 0.118)) * 72)
    TABLE_ROW_HEIGHT_SUBTOTAL = Pt(float(row_heights_config.get("subtotal_inches", 0.139)) * 72)
    TABLE_COLUMN_WIDTHS = [
        Inches(0.65),
        Inches(0.50),
        Inches(0.35),
        Inches(0.43),
        Inches(0.35),
        Inches(0.40),
        Inches(0.72),
    ] + [Inches(0.375)] * 16
    TABLE_TOP_OVERRIDE = Inches(float(table_position_config.get("top_inches", 0.812)))

    TABLE_CELL_STYLE_CONTEXT = CellStyleContext(
        margin_left_right_pt=3.6,
        margin_emu_lr=MARGIN_EMU_LR,
        default_font_name=DEFAULT_FONT_NAME,
        font_size_header=FONT_SIZE_HEADER,
        font_size_body=FONT_SIZE_BODY,
        color_black=CLR_BLACK,
        color_light_gray_text=CLR_LIGHT_GRAY_TEXT,
        color_table_gray=CLR_TABLE_GRAY,
        color_header_green=CLR_HEADER_GREEN,
        color_subtotal_gray=CLR_SUBTOTAL_GRAY,
        color_tv=CLR_TELEVISION,
        color_digital=CLR_DIGITAL,
        color_ooh=CLR_OOH,
        color_other=CLR_OTHER,
    )

    CHART_STYLE_CONTEXT = ChartStyleContext(
        font_name=DEFAULT_FONT_NAME,
        title_font_size=FONT_SIZE_CHART_TITLE,
        label_font_size=FONT_SIZE_CHART_LABELS,
        font_color=CLR_BLACK,
    )

    CHART_COLOR_MAPPING = {
        "Television": CLR_TELEVISION,
        "Digital": CLR_DIGITAL,
        "OOH": CLR_OOH,
        "Other": CLR_OTHER,
        "Awareness": CLR_TELEVISION,
        "Consideration": CLR_BLACK,
        "Purchase": CLR_OOH,
        "Always On": CLR_TELEVISION,
        "Brand": CLR_DIGITAL,
        "Product": CLR_OOH,
    }

    CHART_COLOR_CYCLE = [CLR_TELEVISION, CLR_DIGITAL, CLR_OOH, CLR_OTHER]

    MAX_ROWS_PER_SLIDE = int(table_config.get("max_rows_per_slide", 17))
    SPLIT_STRATEGY = table_config.get("split_strategy", "by_campaign")
    SHOW_CHARTS_ON_SPLITS = table_config.get("show_charts_on_splits", "all")
    SHOW_CARRIED_SUBTOTAL = bool(table_config.get("show_carried_subtotal", True))
    CONTINUATION_INDICATOR = table_config.get("continuation_indicator", " (Continued)")

    ELEMENT_COORDINATES = {
        "title": _coord_from_config(
            title_position_config,
            {"left": 0.184, "top": 0.308, "width": 2.952, "height": 0.370},
        ),
        "main_table": _coord_from_config(
            table_position_config,
            {"left": 0.184, "top": 0.812, "width": 9.299, "height": 2.338},
        ),
        "comments_title": _coord_from_config(
            comments_title_pos_config,
            {"left": 1.097, "top": 3.697, "width": 0.640, "height": 0.151},
        ),
        "comments_box": _coord_from_config(
            comments_box_pos_config,
            {"left": 0.184, "top": 3.886, "width": 2.466, "height": 1.489},
        ),
        "chart_1": _coord_from_config(
            chart_positions_config.get("funnel_chart", {}),
            {"left": 2.650, "top": 3.300, "width": 2.466, "height": 2.000},
        ),
        "chart_2": _coord_from_config(
            chart_positions_config.get("media_type_chart", {}),
            {"left": 4.725, "top": 3.300, "width": 2.647, "height": 2.000},
        ),
        "chart_3": _coord_from_config(
            chart_positions_config.get("campaign_type_chart", {}),
            {"left": 6.985, "top": 3.300, "width": 2.647, "height": 2.000},
        ),
        "tv_legend_color": {"left": 6.645, "top": 0.438, "width": 0.259, "height": 0.139},
        "tv_legend_text": {"left": 6.841, "top": 0.416, "width": 0.612, "height": 0.219},
        "digital_legend_color": {"left": 7.463, "top": 0.449, "width": 0.259, "height": 0.139},
        "digital_legend_text": {"left": 7.658, "top": 0.416, "width": 0.467, "height": 0.219},
        "ooh_legend_color": {"left": 8.196, "top": 0.449, "width": 0.259, "height": 0.139},
        "ooh_legend_text": {"left": 8.392, "top": 0.416, "width": 0.393, "height": 0.219},
        "other_legend_color": {"left": 8.866, "top": 0.449, "width": 0.259, "height": 0.139},
        "other_legend_text": {"left": 9.061, "top": 0.416, "width": 0.439, "height": 0.219},
    }


def configure(config: Config) -> None:
    _initialize_from_config(config)


_initialize_from_config(load_master_config())

# Define the constant we need (EXACTLY = 1 is the standard value)
class WD_ROW_HEIGHT_RULE:
    AT_LEAST = 0  # Add this for more flexible row height
    EXACTLY = 1

TABLE_HEIGHT_RULE_AVAILABLE = True

# --- Constants for Named Shapes (from Template_Refactoring_Guide.md) ---
SHAPE_NAME_TITLE = "TitlePlaceholder"
SHAPE_NAME_TABLE = "MainDataTable"
SHAPE_NAME_COMMENTS_TITLE = "CommentsTitle"
SHAPE_NAME_COMMENTS_BOX = "CommentsBox"
SHAPE_NAME_FUNNEL_CHART = "FunnelChart"
SHAPE_NAME_MEDIA_TYPE_CHART = "MediaTypeChart"
SHAPE_NAME_CAMPAIGN_TYPE_CHART = "CampaignTypeChart"
SHAPE_NAME_TV_LEGEND_COLOR = "TelevisionLegendColor"
SHAPE_NAME_TV_LEGEND_TEXT = "TelevisionLegendText"
SHAPE_NAME_DIGITAL_LEGEND_COLOR = "DigitalLegendColor"
SHAPE_NAME_DIGITAL_LEGEND_TEXT = "DigitalLegendText"
SHAPE_NAME_OOH_LEGEND_COLOR = "OOHLegendColor"
SHAPE_NAME_OOH_LEGEND_TEXT = "OOHLegendText"
SHAPE_NAME_OTHER_LEGEND_COLOR = "OtherLegendColor"
SHAPE_NAME_OTHER_LEGEND_TEXT = "OtherLegendText"

# Must exactly match the placeholder name in Slide Master
TABLE_PLACEHOLDER_NAME = "Table Placeholder 1"

# --- Color Constants (QA checklist verified RGB values) ---
CLR_BLACK = RGBColor(0, 0, 0)
CLR_WHITE = RGBColor(255, 255, 255)
CLR_LIGHT_GRAY_TEXT = _rgb_color(ui_colors_config.get("light_gray_text", {}).get("rgb"), (191, 191, 191))
CLR_TABLE_GRAY = _rgb_color(ui_colors_config.get("table_gray", {}).get("rgb"), (191, 191, 191))
CLR_HEADER_GREEN = _rgb_color(ui_colors_config.get("header_green", {}).get("rgb"), (56, 236, 4))
CLR_COMMENTS_GRAY = _rgb_color(ui_colors_config.get("comments_gray", {}).get("rgb"), (242, 242, 242))
CLR_SUBTOTAL_GRAY = _rgb_color(ui_colors_config.get("subtotal_gray", {}).get("rgb"), (217, 217, 217))

CLR_TELEVISION = _rgb_color(media_colors_config.get("television", {}).get("rgb"), (113, 212, 141))
CLR_DIGITAL = _rgb_color(media_colors_config.get("digital", {}).get("rgb"), (253, 242, 183))
CLR_OOH = _rgb_color(media_colors_config.get("ooh", {}).get("rgb"), (255, 191, 0))
CLR_OTHER = _rgb_color(media_colors_config.get("other", {}).get("rgb"), (176, 211, 255))

# --- PIXEL-PERFECT FONT CONSTANTS ---
DEFAULT_FONT_NAME = fonts_config.get("default_family", "Calibri")
FONT_SIZE_HEADER = Pt(float(font_sizes.get("header", 7.5)))
FONT_SIZE_BODY = Pt(float(font_sizes.get("body", 7.0)))
FONT_SIZE_CHART_TITLE = Pt(float(font_sizes.get("chart_title", 8.0)))
FONT_SIZE_CHART_LABELS = Pt(float(font_sizes.get("chart_labels", 6.0)))

TABLE_ROW_HEIGHT_HEADER = Pt(float(row_heights_config.get("header_inches", 0.139)) * 72)
TABLE_ROW_HEIGHT_BODY = Pt(float(row_heights_config.get("body_inches", 0.118)) * 72)
TABLE_ROW_HEIGHT_SUBTOTAL = Pt(float(row_heights_config.get("subtotal_inches", 0.139)) * 72)
TABLE_COLUMN_WIDTHS = [
    Inches(0.65),
    Inches(0.50),
    Inches(0.35),
    Inches(0.43),
    Inches(0.35),
    Inches(0.40),
    Inches(0.72),
] + [Inches(0.375)] * 16
TABLE_TOP_OVERRIDE = Inches(float(table_position_config.get("top_inches", 0.812)))

TABLE_CELL_STYLE_CONTEXT = CellStyleContext(
    margin_left_right_pt=3.6,
    margin_emu_lr=MARGIN_EMU_LR,
    default_font_name=DEFAULT_FONT_NAME,
    font_size_header=FONT_SIZE_HEADER,
    font_size_body=FONT_SIZE_BODY,
    color_black=CLR_BLACK,
    color_light_gray_text=CLR_LIGHT_GRAY_TEXT,
    color_table_gray=CLR_TABLE_GRAY,
    color_header_green=CLR_HEADER_GREEN,
    color_subtotal_gray=CLR_SUBTOTAL_GRAY,
    color_tv=CLR_TELEVISION,
    color_digital=CLR_DIGITAL,
    color_ooh=CLR_OOH,
    color_other=CLR_OTHER,
)

CHART_STYLE_CONTEXT = ChartStyleContext(
    font_name=DEFAULT_FONT_NAME,
    title_font_size=FONT_SIZE_CHART_TITLE,
    label_font_size=FONT_SIZE_CHART_LABELS,
    font_color=CLR_BLACK,
)

CHART_COLOR_MAPPING = {
    "Television": CLR_TELEVISION,
    "Digital": CLR_DIGITAL,
    "OOH": CLR_OOH,
    "Other": CLR_OTHER,
    "Awareness": CLR_TELEVISION,
    "Consideration": CLR_BLACK,
    "Purchase": CLR_OOH,
    "Always On": CLR_TELEVISION,
    "Brand": CLR_DIGITAL,
    "Product": CLR_OOH,
}

CHART_COLOR_CYCLE = [CLR_TELEVISION, CLR_DIGITAL, CLR_OOH, CLR_OTHER]

FONT_SIZE_TITLE = Pt(float(font_sizes.get("title", 11.0)))
FONT_SIZE_LEGEND = Pt(float(font_sizes.get("legend", 6.0)))
FONT_SIZE_COMMENTS = Pt(float(font_sizes.get("comments", 9.0)))

# --- TABLE SPLITTING CONSTANTS ---
# Calculate based on available height: 2.34" total height, with header at 0.139" and body rows at 0.118"
# Available for body rows: 2.34 - 0.139 = 2.201"
# Number of body rows that fit: 2.201 / 0.118 = ~18.6, so 18 body rows + 1 header = 19 total
# Reduced to 17 to ensure proper fit with margins and prevent overflow
MAX_ROWS_PER_SLIDE = int(table_config.get("max_rows_per_slide", 17))
SPLIT_STRATEGY = table_config.get("split_strategy", "by_campaign")
SHOW_CHARTS_ON_SPLITS = table_config.get("show_charts_on_splits", "all")
SHOW_CARRIED_SUBTOTAL = bool(table_config.get("show_carried_subtotal", True))
CONTINUATION_INDICATOR = table_config.get("continuation_indicator", " (Continued)")

# --- GEOSPATIAL 2D COORDINATE SYSTEM ---
# Precise positioning for 10" × 5.625" PowerPoint slide (FINAL QA VERIFIED)
# Canvas: 16:9 aspect ratio with origin (0,0) at top-left corner
# All coordinates verified against "Egypt - Centrum" reference slide + Final QA checklist
ELEMENT_COORDINATES = {
    "title": _coord_from_config(
        title_position_config,
        {"left": 0.184, "top": 0.308, "width": 2.952, "height": 0.370},
    ),
    "main_table": _coord_from_config(
        table_position_config,
        {"left": 0.184, "top": 0.812, "width": 9.299, "height": 2.338},
    ),
    "comments_title": _coord_from_config(
        comments_title_pos_config,
        {"left": 1.097, "top": 3.697, "width": 0.640, "height": 0.151},
    ),
    "comments_box": _coord_from_config(
        comments_box_pos_config,
        {"left": 0.184, "top": 3.886, "width": 2.466, "height": 1.489},
    ),
    "chart_1": _coord_from_config(
        chart_positions_config.get("funnel_chart", {}),
        {"left": 2.650, "top": 3.300, "width": 2.466, "height": 2.000},
    ),
    "chart_2": _coord_from_config(
        chart_positions_config.get("media_type_chart", {}),
        {"left": 4.725, "top": 3.300, "width": 2.647, "height": 2.000},
    ),
    "chart_3": _coord_from_config(
        chart_positions_config.get("campaign_type_chart", {}),
        {"left": 6.985, "top": 3.300, "width": 2.647, "height": 2.000},
    ),
    "tv_legend_color": {"left": 6.645, "top": 0.438, "width": 0.259, "height": 0.139},
    "tv_legend_text": {"left": 6.841, "top": 0.416, "width": 0.612, "height": 0.219},
    "digital_legend_color": {"left": 7.463, "top": 0.449, "width": 0.259, "height": 0.139},
    "digital_legend_text": {"left": 7.658, "top": 0.416, "width": 0.467, "height": 0.219},
    "ooh_legend_color": {"left": 8.196, "top": 0.449, "width": 0.259, "height": 0.139},
    "ooh_legend_text": {"left": 8.392, "top": 0.416, "width": 0.393, "height": 0.219},
    "other_legend_color": {"left": 8.866, "top": 0.449, "width": 0.259, "height": 0.139},
    "other_legend_text": {"left": 9.061, "top": 0.416, "width": 0.439, "height": 0.219},
}

def get_element_position(element_name):
    """
    Get positioning coordinates for a named element in inches.
    
    Args:
        element_name: Key from ELEMENT_COORDINATES dictionary
        
    Returns:
        dict: Dictionary with left, top, width, height in Inches objects
    """
    if element_name not in ELEMENT_COORDINATES:
        logger.error(f"Element '{element_name}' not found in coordinate system")
        return None
    
    coords = ELEMENT_COORDINATES[element_name]
    return {
        'left': Inches(coords['left']),
        'top': Inches(coords['top']), 
        'width': Inches(coords['width']),
        'height': Inches(coords['height'])
    }

logger = logging.getLogger("amp_automation.legacy")

# --- CORRECTED Excel Column Indices for YOUR ACTUAL FILE ---
# Updated to match YOUR BULK_PLAN_EXPORT_2025_08_25.xlsx structure
COLUMN_GEOGRAPHY = 10              # Plan - Geography (extract country from hierarchy) - NO CHANGE
COLUMN_GLOBAL_MASTERBRAND = 17     # Plan - Brand - NO CHANGE
COLUMN_MEDIA_TYPE = 20             # Media Type - NO CHANGE
COLUMN_CAMPAIGN_NAMES = 83         # **Campaign Name(s) - WAS 85
COLUMN_CAMPAIGN_TYPE = 84          # **Campaign Type - WAS 86
COLUMN_FUNNEL_STAGE = 95           # **Funnel Stage - WAS 97
COLUMN_YEAR = 15                   # Plan - Year - NO CHANGE
COLUMN_FLIGHT_START_DATE = 110     # **Flight Start Date (for month extraction) - NO CHANGE
COLUMN_NET_COST = 71               # *Net Cost (budget data) - WAS 73
COLUMN_NATIONAL_GRP = 55           # GRP - WAS 56

# Additional Lumina columns for enhanced TV data (CORRECTED)
COLUMN_REACH_1PLUS = 104           # Reach 1+ - WAS 106
COLUMN_REACH_3PLUS = 64            # Reach 3+ - NO CHANGE
COLUMN_FREQUENCY = 105             # Frequency - WAS 107

# Legacy column references (kept for backwards compatibility)
COLUMN_COUNTRY = 10                # Now uses COLUMN_GEOGRAPHY
COLUMN_MONTH_START_DATE = 110      # Now uses COLUMN_FLIGHT_START_DATE

# Actual column name for GRP data, used if reading by name
COLUMN_NATIONAL_GRP_NAME = 'National GRP [Current]'

# --- Data Processing Constants (from v1.5) ---
MEDIA_TYPE_MAPPING = {
    'Television': 'TV',
    'TV': 'TV',
    'Digital': 'Digital',
    'OOH': 'OOH',
    'Other': 'Other',
    'Print': 'Other', # Group Print under Other
    'Radio': 'Other', # Group Radio under Other
    'Cinema': 'Other' # Group Cinema under Other
}

MONTH_MAP = {
    1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
    7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
}

EXPECTED_COLUMNS = [
    COLUMN_GLOBAL_MASTERBRAND,
    COLUMN_MEDIA_TYPE,
    COLUMN_CAMPAIGN_NAMES,
    COLUMN_COUNTRY,
    COLUMN_NET_COST,
    COLUMN_MONTH_START_DATE # Used to derive month
]

# --- Helper Functions ---
def _get_shape_by_name(slide, name):
    """Find a shape on a slide by its name."""
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    logger.warning(f"Shape with name '{name}' not found on slide.")
    return None

def _copy_text_box(source_shape, target_slide, new_name=None, new_text=None):
    """Copies a source text box shape to a target slide, with optional new name and text."""
    if not source_shape.has_text_frame:
        logger.warning(f"Source shape '{source_shape.name}' does not have a text frame. Cannot copy as text box.")
        return None

    left, top, width, height = source_shape.left, source_shape.top, source_shape.width, source_shape.height
    new_shape = target_slide.shapes.add_textbox(left, top, width, height)
    new_shape.name = new_name if new_name else source_shape.name

    # Copy text frame properties
    source_tf = source_shape.text_frame
    new_tf = new_shape.text_frame

    new_tf.word_wrap = source_tf.word_wrap
    # MSO_AUTO_SIZE enum: NONE = 0, SHAPE_TO_FIT_TEXT = 1, TEXT_TO_FIT_SHAPE = 2
    # Ensure auto_size is handled correctly; direct copy might be problematic if enum values differ or not available.
    # For simplicity, let's try direct copy if it's simple. Otherwise, might need specific handling.
    try:
        new_tf.auto_size = source_tf.auto_size
    except AttributeError:
        logger.debug(f"Could not directly copy auto_size for shape '{new_shape.name}'. Defaulting.")
        # Default or no action if direct copy fails

    if new_text is not None:
        new_tf.text = new_text
    else:
        new_tf.text = source_tf.text

    # Copy paragraph formatting for the first paragraph (if text exists)
    if source_tf.paragraphs and new_tf.paragraphs:
        source_para = source_tf.paragraphs[0]
        new_para = new_tf.paragraphs[0]
        new_para.alignment = source_para.alignment
        
        if source_para.runs and new_para.runs:
            source_run = source_para.runs[0]
            new_run = new_para.runs[0]
            new_run.font.name = source_run.font.name
            new_run.font.size = source_run.font.size
            if source_run.font.color.type == MSO_COLOR_TYPE.RGB:
                 if hasattr(source_run.font.color, 'rgb'):
                    new_run.font.color.rgb = source_run.font.color.rgb
            new_run.font.bold = source_run.font.bold
            new_run.font.italic = source_run.font.italic
        elif not new_para.runs and new_text == "": # Handle empty text box case for font styling
            # If new_text is empty, add a run to apply font style from source
            if source_para.runs:
                source_run = source_para.runs[0]
                new_run = new_para.add_run()
                new_run.font.name = source_run.font.name
                new_run.font.size = source_run.font.size
                if source_run.font.color.type == MSO_COLOR_TYPE.RGB:
                     if hasattr(source_run.font.color, 'rgb'):
                        new_run.font.color.rgb = source_run.font.color.rgb
                new_run.font.bold = source_run.font.bold
                new_run.font.italic = source_run.font.italic
                new_para.text = "" # Clear the run's default text if any

    logger.debug(f"Copied text box '{source_shape.name}' to '{new_shape.name}' on target slide.")
    return new_shape

def _copy_shape(source_shape, target_slide, new_name=None):
    """Copies a basic source auto shape to a target slide, with optional new name."""
    if source_shape.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
        logger.warning(f"Source shape '{source_shape.name}' is not an AUTO_SHAPE (type: {source_shape.shape_type}). Skipping copy.")
        return None

    left, top, width, height = source_shape.left, source_shape.top, source_shape.width, source_shape.height
    # Add the shape with the same auto_shape_type
    new_shape = target_slide.shapes.add_shape(
        source_shape.auto_shape_type, left, top, width, height
    )
    new_shape.name = new_name if new_name else source_shape.name

    # Copy fill properties
    source_fill = source_shape.fill
    new_fill = new_shape.fill
    # Assuming solid fill for legend color boxes
    new_fill.solid()
    if hasattr(source_fill.fore_color, 'rgb'):
        new_fill.fore_color.rgb = source_fill.fore_color.rgb
    else:
        logger.debug(f"Source shape '{source_shape.name}' fill fore_color has no rgb. Defaulting fill.")

    # Copy line properties
    source_line = source_shape.line
    new_line = new_shape.line
    if hasattr(source_line.color, 'rgb') and source_line.color.type != 0: # Check if color is set and not NO_FILL implicitly by type
        new_line.color.rgb = source_line.color.rgb
    elif source_line.fill.type == 0: # NO_FILL (0)
        new_line.fill.background() # Or new_line.fill.solid() with transparency, or just leave as default if that's no line
        logger.debug("Source shape '{source_shape.name}' line has NO_FILL. Applying background fill to line of new shape.")
    else:
        logger.debug(f"Source shape '{source_shape.name}' line color has no rgb or complex fill. Defaulting line color.")
    
    new_line.width = source_line.width # Width is in EMUs, direct copy is fine

    logger.debug(f"Copied shape '{source_shape.name}' to '{new_shape.name}' on target slide.")
    return new_shape

def is_empty_formatted_value(formatted_value):
    """Check if a formatted value represents empty/zero data that should not be displayed"""
    if not formatted_value:
        return True
    
    # Standard empty indicators
    if formatted_value in ["-", "", "0.0%"]:
        return True
    
    # Budget-specific empty indicators
    if formatted_value in ["£0K", "£0", "£0.00K"]:
        return True
        
    return False

def format_number(value, is_budget=False, is_percentage=False, is_grp=False, is_monthly_column=False):
    """Formats numbers for display in tables, with improved accuracy for budgets and enhanced zero handling."""
    if pd.isna(value):
        if is_percentage:
            return "0.0%"
        return ""  # Return empty string for NaN values to trigger "-" display
    
    numeric_value = 0
    try:
        numeric_value = float(value)
    except ValueError:
        if is_percentage:
            return "0.0%"
        return ""  # Return empty string for invalid values to trigger "-" display

    if numeric_value == 0 or abs(numeric_value) < ZERO_THRESHOLD:  # Enhanced zero detection including near-zero values
        if is_percentage:
            return "0.0%"
        # For budget and GRP values, return empty string to trigger "-" display in cell styling
        if is_budget or is_grp:
            return ""  # This will trigger "-" display in _apply_table_cell_styling
        return ""

    # Handle percentages first
    if is_percentage:
        return f"{numeric_value:.1f}%"

    # Handle GRPs with K suffix for thousands
    if is_grp:
        abs_value = abs(numeric_value)
        if abs_value >= 1_000:
            formatted_val = numeric_value / 1_000.0
            # Show one decimal place for accuracy (e.g., 23.5K)
            return f"{formatted_val:.1f}K"
        else:
            # For values less than 1000, show as integer
            return f"{int(numeric_value)}"

    # IMPROVED BUDGET FORMATTING: More accurate representation
    if is_budget:
        abs_value = abs(numeric_value)
        
        # For values >= 1M, use millions
        if abs_value >= 1_000_000:
            formatted_val = numeric_value / 1_000_000.0
            # For monthly columns, always use whole numbers (no decimals)
            if is_monthly_column:
                return f"£{formatted_val:.0f}M"
            else:
                # Use 1 decimal place for millions to preserve accuracy in main Budget column
                if formatted_val == int(formatted_val):
                    return f"£{formatted_val:.0f}M"
                else:
                    return f"£{formatted_val:.1f}M"
        
        # For values >= 1K, use thousands
        elif abs_value >= 1_000:
            formatted_val = numeric_value / 1_000.0
            # For monthly columns, always use whole numbers (no decimals)
            if is_monthly_column:
                return f"£{formatted_val:.0f}K"
            else:
                # Use 1 decimal place for thousands when needed for accuracy in main Budget column
                if formatted_val == int(formatted_val):
                    return f"£{formatted_val:.0f}K"
                else:
                    return f"£{formatted_val:.1f}K"
        
        # For values < 1K, show in pounds with K suffix for consistency
        else:
            # CRITICAL FIX: For monthly columns with values < 1K, show actual amount in pounds
            # This prevents £500 from displaying as "£0K" which appears empty
            if is_monthly_column and abs_value > 0:
                return f"£{int(numeric_value)}"
            
            formatted_val = numeric_value / 1_000.0
            # For monthly columns, always round to whole K units
            if is_monthly_column:
                return f"£{formatted_val:.0f}K"
            else:
                return f"£{formatted_val:.2f}K"
    
    # Non-budget formatting (fallback)
    abs_value = abs(numeric_value)
    if abs_value >= 1_000_000:
        suffix = 'M'
        divisor = 1_000_000.0
    elif abs_value >= 1_000:
        suffix = 'K'
        divisor = 1_000.0
    else:
        suffix = 'K'
        divisor = 1_000.0

    formatted_val = numeric_value / divisor
    formatted_str = f"{formatted_val:.1f}{suffix}"
    return formatted_str

def load_and_prepare_data(excel_path):  # backward compatibility proxy
    active_logger = logger if 'logger' in globals() and logger is not None else logging.getLogger("amp_automation.data")
    data_set = modular_load_and_prepare_data(excel_path, MASTER_CONFIG, active_logger)
    return data_set.frame

def _prepare_main_table_data_detailed(df, region, masterbrand, year=None, excel_path=None):
    """
    Prepare table data for a specific region/masterbrand combination.
    Adapted from v1.5 create_campaign_table function for v2.0 template approach.
    
    Args:
        df: DataFrame with loaded Excel data
        region: Region name to filter by
        masterbrand: Masterbrand name to filter by
        
    Returns:
        tuple: (table_data, cell_metadata) where:
            - table_data: List of lists representing table rows
            - cell_metadata: Dictionary with metadata for each cell for conditional formatting
        or (None, None) if no data found
    """
    try:
        year_text = f" - {year}" if year is not None else ""
        logger.info(f"Preparing table data for {region} - {masterbrand}{year_text}")
        
        # **ENHANCED DEBUGGING**: Log data filtering process
        logger.debug(f"Total rows in DataFrame before filtering: {len(df)}")
        logger.debug(f"Unique countries in data: {sorted(df['Country'].dropna().unique())}")
        logger.debug(f"Unique masterbrands in data: {sorted(df['Brand'].dropna().unique())}")
        if year is not None:
            logger.debug(f"Unique years in data: {sorted(df['Year'].dropna().unique())}")
        
        # Month mapping for quarterly calculations
        month_mapping = {
            'January': 1, 'Jan': 1, 'JAN': 1, '1': 1, 1: 1,
            'February': 2, 'Feb': 2, 'FEB': 2, '2': 2, 2: 2,
            'March': 3, 'Mar': 3, 'MAR': 3, '3': 3, 3: 3,
            'April': 4, 'Apr': 4, 'APR': 4, '4': 4, 4: 4,
            'May': 5, 'MAY': 5, '5': 5, 5: 5,
            'June': 6, 'Jun': 6, 'JUN': 6, '6': 6, 6: 6,
            'July': 7, 'Jul': 7, 'JUL': 7, '7': 7, 7: 7,
            'August': 8, 'Aug': 8, 'AUG': 8, '8': 8, 8: 8,
            'September': 9, 'Sep': 9, 'SEP': 9, '9': 9, 9: 9,
            'October': 10, 'Oct': 10, 'OCT': 10, '10': 10, 10: 10,
            'November': 11, 'Nov': 11, 'NOV': 11, '11': 11, 11: 11,
            'December': 12, 'Dec': 12, 'DEC': 12, '12': 12
        }
        
        # Constants for table structure
        MONTHS = ['JAN', 'FEB', 'MAR', 'Q1', 'APR', 'MAY', 'JUN', 'Q2', 'JUL', 'AUG', 'SEP', 'Q3', 'OCT', 'NOV', 'DEC', 'Q4']
        GRP_DATA_IDENTIFIER = "GRPs"  # Identifier for GRP sub-rows
        
        # Filter data for the specific region, masterbrand, and year
        # Use the processed column names from load_and_prepare_data
        filter_conditions = [
            (df['Country'].astype(str).str.strip() == region),
            (df['Brand'].astype(str).str.strip() == masterbrand)
        ]
        
        if year is not None:
            filter_conditions.append(df['Year'].astype(str).str.strip() == str(year))
        
        filtered_df = df[
            filter_conditions[0] & filter_conditions[1] & (filter_conditions[2] if len(filter_conditions) > 2 else True)
        ].copy()
        
        # **ENHANCED DEBUGGING**: Log filtering results
        logger.debug(f"Rows after filtering for {region} - {masterbrand}: {len(filtered_df)}")
        
        if filtered_df.empty:
            logger.warning(f"No data found for {region} - {masterbrand}")
            # **ENHANCED DEBUGGING**: Show available combinations for troubleshooting
            available_combinations = df.groupby(['Country', 'Brand']).size().reset_index(name='count')
            logger.debug(f"Available country/masterbrand combinations:")
            for _, combo in available_combinations.iterrows():
                logger.debug(f"  {combo['Country']} - {combo['Brand']} ({combo['count']} rows)")
            return None, None
            
        logger.info(f"Found {len(filtered_df)} rows for {region} - {masterbrand}")
        
        # Get unique campaigns and sort them
        campaign_column = 'Campaign Name'
        campaigns = sorted(filtered_df[campaign_column].dropna().unique())
        
        # **ENHANCED DEBUGGING**: Log campaign details
        logger.info(f"Found {len(campaigns)} unique campaigns for {region} - {masterbrand}:")
        for i, campaign in enumerate(campaigns):
            campaign_rows = len(filtered_df[filtered_df[campaign_column] == campaign])
            logger.info(f"  {i+1}. '{campaign}' ({campaign_rows} rows)")
        
        # **ENHANCED DEBUGGING**: Check for potential data quality issues
        if any(campaign in str(campaigns) for campaign in ['Gender', 'You Did It']):
            logger.warning(f"DATA QUALITY ALERT: Suspicious campaigns found for {region} - {masterbrand}: {campaigns}")
            logger.warning("This may indicate incorrect data filtering or source data issues")
        
        # Initialize table data with headers
        table_data = [['CAMPAIGN NAME', 'BUDGET', 'GRPs', 'REACH', 'OTS', '%', 'MEDIA TYPE'] + MONTHS]
        monthly_totals = [0.0] * len(MONTHS)
        grp_monthly_totals = [0.0] * len(MONTHS)
        
        # Initialize cell metadata for conditional formatting
        # Structure: {(row_idx, col_idx): {'has_data': bool, 'media_type': str, 'value': float}}
        cell_metadata = {}
        
        # Calculate total budget for percentage calculations
        total_budget_for_percentage = filtered_df['Total Cost'].sum()
        grand_total_grp = 0  # Sum of campaign total GRPs for footer
        
        logger.debug(f"Total budget for percentage calculation: £{total_budget_for_percentage:,.2f}")
        
        for campaign_name in campaigns:
            if pd.isna(campaign_name) or str(campaign_name).strip() == '':
                continue
            
            campaign_df = filtered_df[filtered_df[campaign_column] == campaign_name]
            logger.debug(f"Processing campaign: '{campaign_name}' ({len(campaign_df)} rows)")
            
            campaign_total_budget = campaign_df['Total Cost'].sum()
            campaign_percentage = (campaign_total_budget / total_budget_for_percentage * 100) if total_budget_for_percentage > 0 else 0
            
            logger.debug(f"  Campaign budget: £{campaign_total_budget:,.2f} ({campaign_percentage:.1f}%)")
            
            # Calculate total TV GRP for this campaign
            campaign_total_tv_grp_value = 0
            if 'GRP' in campaign_df.columns:
                campaign_tv_df = campaign_df[campaign_df['Media Type'].astype(str).str.upper().isin(['TV', 'TELEVISION'])]
                if not campaign_tv_df.empty:
                    grp_sum = campaign_tv_df['GRP'].sum()
                    if not pd.isna(grp_sum):
                        campaign_total_tv_grp_value = grp_sum
            
            logger.debug(f"  Campaign TV GRPs: {campaign_total_tv_grp_value}")
            
            # Get unique media types for this campaign, prioritizing TV first
            media_types_for_campaign = sorted(campaign_df['Media Type'].fillna('').astype(str).unique())
            media_types_for_campaign.sort(key=lambda mt: 0 if normalize_media_type(str(mt)) == 'Television' else 1)
            
            logger.debug(f"  Media types: {media_types_for_campaign}")
            
            first_media_type_for_campaign = True
            
            for media_type_original in media_types_for_campaign:
                if pd.isna(media_type_original) or str(media_type_original).strip() == '':
                    continue
                
                current_media_type_normalized = normalize_media_type(str(media_type_original))
                media_df = campaign_df[campaign_df['Media Type'] == media_type_original]
                
                monthly_budget_data_formatted = ['-'] * len(MONTHS)
                monthly_grp_data_formatted = ['-'] * len(MONTHS)
                
                is_tv_current_media = current_media_type_normalized == 'Television'
                
                # For wide-format data, monthly budgets are already in columns
                # No need to group by Month column
                
                # Populate monthly budget data with quarterly totals
                q_budget_total = 0
                current_row_idx = len(table_data)  # Track current row index for metadata
                
                for i, month_name in enumerate(MONTHS):
                    month_col_idx = i + 7  # Month columns start after first 7 columns
                    
                    if month_name.startswith('Q'):  # Q1, Q2, Q3, Q4
                        formatted_budget = format_number(q_budget_total, is_budget=True, is_monthly_column=True)
                        monthly_budget_data_formatted[i] = formatted_budget
                        monthly_totals[i] += q_budget_total
                        # Add metadata for quarter columns - FIXED: align with display logic
                        cell_metadata[(current_row_idx, month_col_idx)] = {
                            'has_data': not is_empty_formatted_value(formatted_budget),
                            'media_type': current_media_type_normalized,
                            'value': q_budget_total
                        }
                        q_budget_total = 0
                    else:
                        # For wide-format data, get budget directly from month column
                        # Convert month name to proper case (Jan, Feb, Mar, etc.)
                        month_col_name = month_name.capitalize() if len(month_name) == 3 else month_name
                        if month_col_name in media_df.columns:
                            budget_val = media_df[month_col_name].sum()
                            formatted_budget = format_number(budget_val, is_budget=True, is_monthly_column=True)
                            monthly_budget_data_formatted[i] = formatted_budget
                            monthly_totals[i] += budget_val
                            q_budget_total += budget_val
                            # Add metadata for month columns - FIXED: align with display logic
                            cell_metadata[(current_row_idx, month_col_idx)] = {
                                'has_data': not is_empty_formatted_value(formatted_budget),
                                'media_type': current_media_type_normalized,
                                'value': budget_val
                            }
                        else:
                            logger.warning(f"Month column {month_name} not found in data.")
                            cell_metadata[(current_row_idx, month_col_idx)] = {
                                'has_data': False,
                                'media_type': current_media_type_normalized,
                                'value': 0
                            }
                
                # Process GRP data for TV media types
                if is_tv_current_media and 'GRP' in df.columns:
                    # For wide-format data, GRP is in a single column, not monthly
                    q_grp_total = 0
                    grp_row_idx = current_row_idx + 1  # GRP row comes after main media row
                    
                    for i, month_name in enumerate(MONTHS):
                        month_col_idx = i + 7  # Month columns start after first 7 columns
                        
                        if month_name.startswith('Q'):  # Q1, Q2, Q3, Q4
                            formatted_grp = format_number(q_grp_total, is_grp=True)
                            monthly_grp_data_formatted[i] = formatted_grp
                            grp_monthly_totals[i] += q_grp_total
                            # Add metadata for GRP quarter columns - FIXED: align with display logic
                            cell_metadata[(grp_row_idx, month_col_idx)] = {
                                'has_data': not is_empty_formatted_value(formatted_grp),
                                'media_type': 'GRPs',
                                'value': q_grp_total
                            }
                            q_grp_total = 0
                        else:
                            # Use month-specific aggregation from raw data
                            metrics = get_month_specific_tv_metrics(excel_path, region, masterbrand, campaign_name, year, month_name)
                            grp_val = metrics['grp_sum']
                            
                            
                            if grp_val > 0:
                                formatted_grp = format_number(grp_val, is_grp=True)
                                monthly_grp_data_formatted[i] = formatted_grp
                                grp_monthly_totals[i] += grp_val
                                q_grp_total += grp_val
                            else:
                                formatted_grp = '-'
                                monthly_grp_data_formatted[i] = formatted_grp
                            
                            # Add metadata for GRP month columns - FIXED: align with display logic
                            cell_metadata[(grp_row_idx, month_col_idx)] = {
                                'has_data': not is_empty_formatted_value(formatted_grp),
                                'media_type': 'GRPs',
                                'value': grp_val
                            }
                
                # CRITICAL FIX: Display campaign info for ALL media types, not just first
                # This resolves the issue where Digital campaigns appeared with blank names
                campaign_name_display = str(campaign_name).upper()
                campaign_budget_display = format_number(campaign_total_budget, is_budget=True)
                campaign_percentage_display = format_number(campaign_percentage, is_percentage=True)
                
                # Display total GRP for the campaign only on the primary TV row
                campaign_total_grp_display = format_number(campaign_total_tv_grp_value, is_grp=True) if first_media_type_for_campaign and is_tv_current_media else ""
                if first_media_type_for_campaign and is_tv_current_media:
                    grand_total_grp += campaign_total_tv_grp_value  # Accumulate for footer
                
                # Calculate TOTAL REACH and TOTAL FREQ for TV campaigns
                total_reach_display = ""
                total_freq_display = ""
                if first_media_type_for_campaign and is_tv_current_media:
                    # Calculate TOTAL REACH as average of all Reach@1+ values for this campaign
                    # Get all monthly Reach@1+ values for this campaign from raw data
                    all_reach1_values = []
                    total_grps_for_campaign = 0
                    
                    # Collect all monthly reach@1+ values and GRPs for this campaign
                    for month in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']:
                        metrics = get_month_specific_tv_metrics(excel_path, region, masterbrand, campaign_name, year, month)
                        if not pd.isna(metrics['reach1_avg']) and metrics['reach1_avg'] > 0:
                            all_reach1_values.append(metrics['reach1_avg'])
                        if not pd.isna(metrics['grp_sum']) and metrics['grp_sum'] > 0:
                            total_grps_for_campaign += metrics['grp_sum']
                    
                    # Calculate TOTAL REACH as average of all monthly Reach@1+ values
                    if all_reach1_values:
                        total_reach_avg = sum(all_reach1_values) / len(all_reach1_values)
                        reach_pct = total_reach_avg * 100
                        total_reach_display = f"{reach_pct:.0f}%"
                        
                        # Calculate TOTAL FREQ as Total GRPs ÷ Total Average Reach@1+
                        if total_reach_avg > 0 and total_grps_for_campaign > 0:
                            # Fix frequency calculation: GRPs are already percentage, reach is decimal
                            # Convert reach from decimal to percentage for calculation
                            total_freq = total_grps_for_campaign / (total_reach_avg * 100)
                            # Format as whole number with thousands separator
                            total_freq_display = f"{int(round(total_freq)):,}"
                        else:
                            total_freq_display = "-"
                    else:
                        total_reach_display = "-"
                        total_freq_display = "-"
                else:
                    total_reach_display = "-"
                    total_freq_display = "-"

                # Create primary media type row
                primary_row_data = [
                    campaign_name_display,
                    campaign_budget_display,
                    campaign_total_grp_display,
                    total_reach_display,
                    total_freq_display,
                    campaign_percentage_display,
                    current_media_type_normalized
                ] + monthly_budget_data_formatted
                table_data.append(primary_row_data)
                
                # Add GRP sub-row for TV media types
                if is_tv_current_media:
                    grp_sub_row_data = [
                        "",  # Campaign Name blank
                        "",  # Budget blank
                        "",  # TV GRPs (campaign total) blank
                        "",  # TOTAL REACH blank
                        "",  # TOTAL FREQ blank
                        "",  # % blank
                        GRP_DATA_IDENTIFIER  # Media Type is GRP_DATA_IDENTIFIER
                    ] + monthly_grp_data_formatted
                    table_data.append(grp_sub_row_data)
                    
                    # Add Reach@1+ sub-row for TV campaigns
                    # Calculate month-by-month Reach 1+ values
                    monthly_reach_data = []
                    reach_row_idx = len(table_data)  # Get the row index for metadata
                    
                    # Track quarterly values for average calculation
                    q1_values, q2_values, q3_values, q4_values = [], [], [], []
                    current_quarter_values = q1_values
                    
                    for i, month_name in enumerate(MONTHS):
                        month_col_idx = i + 7  # Month columns start after first 7 columns
                        
                        if month_name.startswith('Q'):  # Q1, Q2, Q3, Q4
                            # Calculate quarterly average from accumulated values
                            if current_quarter_values:
                                quarterly_avg = sum(current_quarter_values) / len(current_quarter_values)
                                # Convert decimal average to percentage (0.70 -> 70%)
                                quarterly_pct = quarterly_avg * 100
                                reach1_formatted = f"{quarterly_pct:.0f}%" if quarterly_pct > 0 else "-"
                            else:
                                reach1_formatted = "-"
                                quarterly_avg = 0
                            
                            monthly_reach_data.append(reach1_formatted)
                            # Add cell metadata for coloring - FIXED: align with display logic
                            cell_metadata[(reach_row_idx, month_col_idx)] = {
                                'has_data': not is_empty_formatted_value(reach1_formatted),
                                'media_type': 'Reach@1+',
                                'value': quarterly_avg
                            }
                            
                            # Switch to next quarter
                            if month_name == 'Q1':
                                current_quarter_values = q2_values
                            elif month_name == 'Q2':
                                current_quarter_values = q3_values
                            elif month_name == 'Q3':
                                current_quarter_values = q4_values
                        else:
                            # Use month-specific aggregation from raw data
                            metrics = get_month_specific_tv_metrics(excel_path, region, masterbrand, campaign_name, year, month_name)
                            reach1_avg = metrics['reach1_avg']
                            
                            if not pd.isna(reach1_avg) and reach1_avg > 0:
                                # Convert decimal to percentage (0.70 -> 70%)
                                reach1_pct = reach1_avg * 100
                                reach1_formatted = f"{reach1_pct:.0f}%"
                                month_reach1_avg = reach1_avg
                                current_quarter_values.append(month_reach1_avg)
                            else:
                                reach1_formatted = "-"
                                month_reach1_avg = 0
                            
                            monthly_reach_data.append(reach1_formatted)
                            # Add cell metadata for coloring - FIXED: align with display logic
                            cell_metadata[(reach_row_idx, month_col_idx)] = {
                                'has_data': not is_empty_formatted_value(reach1_formatted),
                                'media_type': 'Reach@1+',
                                'value': month_reach1_avg
                            }
                    
                    reach_sub_row_data = [
                        "",  # Campaign Name blank
                        "-",  # Budget shows dash
                        "-",  # TV GRPs shows dash
                        "-",  # TOTAL REACH shows dash
                        "-",  # TOTAL FREQ shows dash
                        "-",  # % shows dash
                        "Reach@1+"  # Media Type column
                    ] + monthly_reach_data
                    table_data.append(reach_sub_row_data)
                    
                    # Add OTS@1+ sub-row for TV campaigns
                    # Calculate month-by-month OTS@1+ values (GRPs ÷ Reach@1+)
                    monthly_freq_data = []
                    freq_row_idx = len(table_data)  # Get the row index for metadata
                    
                    # Track quarterly values for average calculation
                    q1_values, q2_values, q3_values, q4_values = [], [], [], []
                    current_quarter_values = q1_values
                    
                    for i, month_name in enumerate(MONTHS):
                        month_col_idx = i + 7  # Month columns start after first 7 columns
                        
                        if month_name.startswith('Q'):  # Q1, Q2, Q3, Q4
                            # Calculate quarterly average from accumulated values
                            if current_quarter_values:
                                quarterly_avg = sum(current_quarter_values) / len(current_quarter_values)
                                freq_formatted = f"{quarterly_avg:.0f}" if quarterly_avg > 0 else "-"
                            else:
                                freq_formatted = "-"
                                quarterly_avg = 0
                            
                            monthly_freq_data.append(freq_formatted)
                            # Add cell metadata for coloring - FIXED: align with display logic
                            cell_metadata[(freq_row_idx, month_col_idx)] = {
                                'has_data': not is_empty_formatted_value(freq_formatted),
                                'media_type': 'OTS@1+',
                                'value': quarterly_avg
                            }
                            
                            # Switch to next quarter
                            if month_name == 'Q1':
                                current_quarter_values = q2_values
                            elif month_name == 'Q2':
                                current_quarter_values = q3_values
                            elif month_name == 'Q3':
                                current_quarter_values = q4_values
                        else:
                            # Use month-specific aggregation from raw data
                            metrics = get_month_specific_tv_metrics(excel_path, region, masterbrand, campaign_name, year, month_name)
                            grp_sum = metrics['grp_sum']
                            reach1_avg = metrics['reach1_avg']
                            
                            # Calculate OTS@1+ = GRPs ÷ Reach@1+
                            if (not pd.isna(grp_sum) and grp_sum > 0 and 
                                not pd.isna(reach1_avg) and reach1_avg > 0):
                                # reach1_avg is in decimal form (0.68 = 68%), so divide by 100
                                ots1_value = grp_sum / (reach1_avg * 100)
                                freq_formatted = f"{ots1_value:.0f}"
                                month_freq_avg = ots1_value
                                current_quarter_values.append(month_freq_avg)
                            else:
                                freq_formatted = "-"
                                month_freq_avg = 0
                            
                            monthly_freq_data.append(freq_formatted)
                            # Add cell metadata for coloring - FIXED: align with display logic
                            cell_metadata[(freq_row_idx, month_col_idx)] = {
                                'has_data': not is_empty_formatted_value(freq_formatted),
                                'media_type': 'OTS@1+',
                                'value': month_freq_avg
                            }
                    
                    freq_sub_row_data = [
                        "",  # Campaign Name blank
                        "-",  # Budget shows dash
                        "-",  # TV GRPs shows dash
                        "-",  # TOTAL REACH shows dash
                        "-",  # TOTAL FREQ shows dash
                        "-",  # % shows dash
                        "OTS@1+"  # Media Type column
                    ] + monthly_freq_data
                    table_data.append(freq_sub_row_data)
                    
                    # Add Reach@3+ sub-row for TV campaigns
                    # Calculate month-by-month Reach 3+ values
                    monthly_reach3_data = []
                    reach3_row_idx = len(table_data)  # Get the row index for metadata
                    
                    # Track quarterly values for average calculation
                    q1_values, q2_values, q3_values, q4_values = [], [], [], []
                    current_quarter_values = q1_values
                    
                    for i, month_name in enumerate(MONTHS):
                        month_col_idx = i + 7  # Month columns start after first 7 columns
                        
                        if month_name.startswith('Q'):  # Q1, Q2, Q3, Q4
                            # Calculate quarterly average from accumulated values
                            if current_quarter_values:
                                quarterly_avg = sum(current_quarter_values) / len(current_quarter_values)
                                # Convert decimal average to percentage (0.70 -> 70%)
                                quarterly_pct = quarterly_avg * 100
                                reach3_formatted = f"{quarterly_pct:.0f}%" if quarterly_pct > 0 else "-"
                            else:
                                reach3_formatted = "-"
                                quarterly_avg = 0
                            
                            monthly_reach3_data.append(reach3_formatted)
                            # Add cell metadata for coloring - FIXED: align with display logic
                            cell_metadata[(reach3_row_idx, month_col_idx)] = {
                                'has_data': not is_empty_formatted_value(reach3_formatted),
                                'media_type': 'Reach@3+',
                                'value': quarterly_avg
                            }
                            
                            # Switch to next quarter
                            if month_name == 'Q1':
                                current_quarter_values = q2_values
                            elif month_name == 'Q2':
                                current_quarter_values = q3_values
                            elif month_name == 'Q3':
                                current_quarter_values = q4_values
                        else:
                            # Use month-specific aggregation from raw data
                            metrics = get_month_specific_tv_metrics(excel_path, region, masterbrand, campaign_name, year, month_name)
                            reach3_avg = metrics['reach3_avg']
                            
                            if not pd.isna(reach3_avg) and reach3_avg > 0:
                                # Convert decimal to percentage (0.70 -> 70%)
                                reach3_pct = reach3_avg * 100
                                reach3_formatted = f"{reach3_pct:.0f}%"
                                month_reach3_avg = reach3_avg
                                current_quarter_values.append(month_reach3_avg)
                            else:
                                reach3_formatted = "-"
                                month_reach3_avg = 0
                            
                            monthly_reach3_data.append(reach3_formatted)
                            # Add cell metadata for coloring - FIXED: align with display logic
                            cell_metadata[(reach3_row_idx, month_col_idx)] = {
                                'has_data': not is_empty_formatted_value(reach3_formatted),
                                'media_type': 'Reach@3+',
                                'value': month_reach3_avg
                            }
                    
                    reach3_sub_row_data = [
                        "",  # Campaign Name blank
                        "-",  # Budget shows dash
                        "-",  # TV GRPs shows dash
                        "-",  # TOTAL REACH shows dash
                        "-",  # TOTAL FREQ shows dash
                        "-",  # % shows dash
                        "Reach@3+"  # Media Type column
                    ] + monthly_reach3_data
                    table_data.append(reach3_sub_row_data)
                    
                    # Add OTS@3+ sub-row for TV campaigns (OTS@3+ = GRPs ÷ Reach@3+)
                    # Calculate month-by-month OTS values
                    monthly_ots_data = []
                    ots_row_idx = len(table_data)  # Get the row index for metadata
                    
                    # Track quarterly values for average calculation
                    q1_values, q2_values, q3_values, q4_values = [], [], [], []
                    current_quarter_values = q1_values
                    
                    for i, month_name in enumerate(MONTHS):
                        month_col_idx = i + 7  # Month columns start after first 7 columns
                        
                        if month_name.startswith('Q'):  # Q1, Q2, Q3, Q4
                            # Calculate quarterly average from accumulated values
                            if current_quarter_values:
                                quarterly_avg = sum(current_quarter_values) / len(current_quarter_values)
                                # OTS is a regular number, not a percentage
                                ots_formatted = f"{quarterly_avg:.0f}" if quarterly_avg > 0 else "-"
                            else:
                                ots_formatted = "-"
                                quarterly_avg = 0
                            
                            monthly_ots_data.append(ots_formatted)
                            # Add cell metadata for coloring - FIXED: align with display logic
                            cell_metadata[(ots_row_idx, month_col_idx)] = {
                                'has_data': not is_empty_formatted_value(ots_formatted),
                                'media_type': 'OTS@3+',
                                'value': quarterly_avg
                            }
                            
                            # Switch to next quarter
                            if month_name == 'Q1':
                                current_quarter_values = q2_values
                            elif month_name == 'Q2':
                                current_quarter_values = q3_values
                            elif month_name == 'Q3':
                                current_quarter_values = q4_values
                        else:
                            # Use month-specific aggregation from raw data
                            metrics = get_month_specific_tv_metrics(excel_path, region, masterbrand, campaign_name, year, month_name)
                            grp_sum = metrics['grp_sum']
                            reach3_avg = metrics['reach3_avg']
                            
                            # Calculate OTS@3+ = GRPs ÷ Reach@3+
                            if (not pd.isna(grp_sum) and grp_sum > 0 and 
                                not pd.isna(reach3_avg) and reach3_avg > 0):
                                # reach3_avg is in decimal form (0.15 = 15%), so divide by 100
                                ots_value = grp_sum / (reach3_avg * 100)
                                # OTS is a regular number, not a percentage
                                ots_formatted = f"{ots_value:.0f}"
                                month_ots_avg = ots_value
                                current_quarter_values.append(month_ots_avg)
                            else:
                                ots_formatted = "-"
                                month_ots_avg = 0
                            
                            monthly_ots_data.append(ots_formatted)
                            # Add cell metadata for coloring - FIXED: align with display logic
                            cell_metadata[(ots_row_idx, month_col_idx)] = {
                                'has_data': not is_empty_formatted_value(ots_formatted),
                                'media_type': 'OTS@3+',
                                'value': month_ots_avg
                            }
                    
                    ots_sub_row_data = [
                        "",  # Campaign Name blank
                        "-",  # Budget shows dash
                        "-",  # TV GRPs shows dash
                        "-",  # TOTAL REACH shows dash
                        "-",  # TOTAL FREQ shows dash
                        "-",  # % shows dash
                        "OTS@3+"  # Media Type column
                    ] + monthly_ots_data
                    table_data.append(ots_sub_row_data)
                
                first_media_type_for_campaign = False
        
        # Add total row
        total_row_budget_values = [format_number(total, is_budget=True, is_monthly_column=True) for total in monthly_totals]
        
        total_row = [
            'TOTAL',
            format_number(total_budget_for_percentage, is_budget=True),
            format_number(grand_total_grp, is_grp=True),  # Sum of all campaign total GRPs for TV
            '-',  # TOTAL REACH blank for total row
            '-',  # TOTAL FREQ blank for total row
            format_number(100, is_percentage=True),
            ''  # Media type blank for total row
        ] + total_row_budget_values
        table_data.append(total_row)
        
        logger.info(f"Table data created successfully for {region} - {masterbrand}. Total rows: {len(table_data)}")
        return table_data, cell_metadata
        
    except Exception as e:
        logger.error(f"Error preparing table data for {region} - {masterbrand}: {str(e)}")
        logger.error(traceback.format_exc())
        return None, None

def _split_table_data_by_campaigns(table_data, cell_metadata):
    """
    Split table data into multiple tables if it exceeds MAX_ROWS_PER_SLIDE.
    Keeps complete campaigns together (including TV sub-rows).
    
    Args:
        table_data: List of lists representing table rows
        cell_metadata: Dictionary with metadata for each cell
        
    Returns:
        List of tuples: [(table_data_split, cell_metadata_split, is_continuation), ...]
    """
    if not table_data or len(table_data) <= MAX_ROWS_PER_SLIDE:
        # No need to split
        return [(table_data, cell_metadata, False)]
    
    logger.info(f"Table has {len(table_data)} rows, splitting needed (max: {MAX_ROWS_PER_SLIDE})")
    
    # Identify header and total row
    header_row = table_data[0]
    total_row = table_data[-1]
    
    # Group campaigns with their sub-rows
    campaigns = []
    current_campaign = []
    current_campaign_name = None
    
    # Track which rows are sub-rows
    sub_row_media_types = ['GRPs', 'Reach@1+', 'OTS@1+', 'Reach@3+', 'OTS@3+']
    
    for i in range(1, len(table_data) - 1):  # Skip header and total
        row = table_data[i]
        media_type = row[6] if len(row) > 6 else ""
        campaign_name = row[0] if row[0] else ""
        
        if media_type in sub_row_media_types:
            # This is a sub-row, add to current campaign
            if current_campaign:  # Make sure we have a campaign to add to
                current_campaign.append((i, row))
        else:
            # This is a main row
            if current_campaign and current_campaign_name:
                # Save previous campaign (including its main row and sub-rows)
                campaigns.append({
                    'name': current_campaign_name,
                    'rows': current_campaign,
                    'start_idx': current_campaign[0][0]
                })
            # Start new campaign with this main row
            current_campaign_name = campaign_name
            current_campaign = [(i, row)]
    
    # Don't forget the last campaign
    if current_campaign and current_campaign_name:
        campaigns.append({
            'name': current_campaign_name,
            'rows': current_campaign,
            'start_idx': current_campaign[0][0]
        })
    
    
    # Now split campaigns into slides
    splits = []
    current_split = [header_row]
    current_split_indices = [0]  # Track original indices for metadata
    current_row_count = 1  # Start with header
    split_number = 0
    
    # Track totals for subtotals
    running_total_budget = 0
    running_total_grp = 0
    subtotal_values = None  # Initialize for carried forward rows
    all_split_rows = []  # Track all rows across splits for final total
    
    for campaign in campaigns:
        campaign_row_count = len(campaign['rows'])
        
        # Check if adding this campaign would exceed the limit
        # Reserve 1 row for subtotal/total
        if current_row_count + campaign_row_count + 1 > MAX_ROWS_PER_SLIDE:
            # Complete current split with total/subtotal
            if split_number > 0 or len(splits) > 0:
                # For intermediate splits, use SUBTOTAL
                subtotal_values = _calculate_subtotal_for_split(current_split[1:])  # Exclude header
                subtotal_row = ['SUBTOTAL'] + subtotal_values
                current_split.append(subtotal_row)
                current_split_indices.append(-1)  # Special index for subtotal
            else:
                # For first split when there will be more, use TOTAL
                total_values = _calculate_subtotal_for_split(current_split[1:])  # Exclude header
                total_row_split = ['TOTAL'] + total_values
                current_split.append(total_row_split)
                current_split_indices.append(-1)  # Special index for total
            
            # Save current split
            split_metadata = _extract_metadata_for_indices(cell_metadata, current_split_indices, len(current_split))
            splits.append((current_split, split_metadata, split_number > 0))
            
            # Start new split
            split_number += 1
            current_split = [header_row]
            current_split_indices = [0]
            current_row_count = 1
            
            # Add carried forward subtotal if configured
            if SHOW_CARRIED_SUBTOTAL and splits and subtotal_values:
                # Make sure subtotal_values has the right number of columns
                if len(subtotal_values) == 22:  # Missing campaign name column
                    carried_row = ['CARRIED FORWARD'] + subtotal_values
                else:
                    logger.warning(f"Unexpected subtotal_values length: {len(subtotal_values)}")
                    carried_row = ['CARRIED FORWARD'] + subtotal_values
                current_split.append(carried_row)
                current_split_indices.append(-2)  # Special index for carried forward
                current_row_count += 1
        
        # Add campaign to current split
        for idx, row in campaign['rows']:
            current_split.append(row)
            current_split_indices.append(idx)
            all_split_rows.append(row)  # Track for final total
        current_row_count += campaign_row_count
    
    # Handle the last split
    if len(splits) > 0:
        # This is a continuation slide, calculate grand total from all rows
        grand_total_values = _calculate_subtotal_for_split(all_split_rows)
        grand_total_row = ['TOTAL'] + grand_total_values
        current_split.append(grand_total_row)
        current_split_indices.append(-1)
    else:
        # This is the only slide, add original total row
        current_split.append(total_row)
        current_split_indices.append(len(table_data) - 1)
    
    # Extract metadata for final split
    split_metadata = _extract_metadata_for_indices(cell_metadata, current_split_indices, len(current_split))
    splits.append((current_split, split_metadata, split_number > 0))
    
    logger.info(f"Table split into {len(splits)} slides")
    return splits


def _calculate_subtotal_for_split(rows):
    """Calculate subtotal values for a split section."""
    # Initialize subtotal values
    subtotal_values = []
    
    # Define sub-row media types to exclude from totals
    sub_row_media_types = ['GRPs', 'Reach@1+', 'OTS@1+', 'Reach@3+', 'OTS@3+']
    
    # Skip first column (campaign name), process budget column
    total_budget = 0
    for row in rows:
        # Skip sub-rows when calculating budget totals
        media_type = row[6] if len(row) > 6 else ""
        if media_type in sub_row_media_types:
            continue
            
        if row[1] and row[1] != '-':
            # Extract numeric value from budget string
            budget_str = row[1].replace('£', '').replace('K', '000').replace('M', '000000').replace(',', '')
            try:
                total_budget += float(budget_str)
            except:
                pass
    
    # Format budget
    if total_budget >= 1000000:
        subtotal_values.append(f"£{total_budget/1000000:.1f}M")
    else:
        subtotal_values.append(f"£{int(total_budget/1000)}K")
    
    # Handle other columns (simplified for now)
    # TV GRPs, TOTAL REACH, TOTAL FREQ, %
    subtotal_values.extend(['-', '-', '-', '-'])
    
    # Media Type column
    subtotal_values.append('')
    
    # Monthly columns - sum up values
    for col_idx in range(7, 23):  # Monthly columns
        monthly_total = 0
        for row in rows:
            # Skip sub-rows when calculating monthly totals
            media_type = row[6] if len(row) > 6 else ""
            if media_type in sub_row_media_types:
                continue
                
            if col_idx < len(row) and row[col_idx] and row[col_idx] != '-':
                try:
                    value_str = row[col_idx].replace('£', '').replace('K', '000').replace('M', '000000').replace(',', '')
                    monthly_total += float(value_str)
                except:
                    pass
        
        if monthly_total > 0:
            if monthly_total >= 1000000:
                subtotal_values.append(f"£{monthly_total/1000000:.1f}M")
            else:
                subtotal_values.append(f"£{int(monthly_total/1000)}K")
        else:
            subtotal_values.append('-')
    
    return subtotal_values


def _extract_metadata_for_indices(original_metadata, indices, new_row_count):
    """Extract metadata for specific row indices and remap to new positions."""
    new_metadata = {}
    
    for new_row_idx, original_idx in enumerate(indices):
        if original_idx < 0:
            # Special indices for subtotal/carried forward rows
            continue
            
        # Copy metadata for this row
        for (orig_row, col), meta in original_metadata.items():
            if orig_row == original_idx:
                new_metadata[(new_row_idx, col)] = meta
    
    return new_metadata


def _prepare_media_type_chart_data_detailed(df, region, masterbrand, year=None):
    """Compatibility wrapper for modular media type chart data preparation."""

    return prepare_media_type_chart_data(df, region, masterbrand, year)

def _prepare_funnel_chart_data_detailed(df, region, masterbrand, year=None):
    """Compatibility wrapper for modular funnel chart data preparation."""

    return prepare_funnel_chart_data(df, region, masterbrand, year)

def _prepare_campaign_type_chart_data(df, region, masterbrand, year=None):
    """Compatibility wrapper for modular campaign type chart data preparation."""

    return prepare_campaign_type_chart_data(df, region, masterbrand, year)

def set_title_text_detailed(title_shape, title_text, template_prs):
    logger.debug(f"Attempting to set title text '{title_text}' for shape '{title_shape.name}' using detailed method.")
    # --- Locate the reference title placeholder and its text_frame from the template --- #
    template_title_ref_shape = None
    # Priority: 1. Named 'TitlePlaceholder' on first slide master
    if template_prs.slide_masters:
        master = template_prs.slide_masters[0]
        for placeholder in master.placeholders:
            if placeholder.name == "TitlePlaceholder":
                template_title_ref_shape = placeholder
                logger.debug(f"Found reference title placeholder '{placeholder.name}' by name in slide master.")
                break
            elif placeholder.is_placeholder and placeholder.placeholder_format.type == PP_PLACEHOLDER.TITLE:
                template_title_ref_shape = placeholder
                logger.debug(f"Found reference title placeholder type '{placeholder.placeholder_format.type}' in slide master.")
                break # Take the first one found
    
    # Priority: 2. Named 'TitlePlaceholder' on first slide layout (if not found in master)
    if not template_title_ref_shape and template_prs.slide_layouts:
        layout = template_prs.slide_layouts[0] # Assuming first layout is relevant
        for placeholder in layout.placeholders:
            if placeholder.name == "TitlePlaceholder":
                template_title_ref_shape = placeholder
                logger.debug(f"Found reference title placeholder '{placeholder.name}' by name in slide layout '{layout.name}'.")
                break
            elif placeholder.is_placeholder and placeholder.placeholder_format.type == PP_PLACEHOLDER.TITLE:
                template_title_ref_shape = placeholder
                logger.debug(f"Found reference title placeholder type '{placeholder.placeholder_format.type}' in slide layout '{layout.name}'.")
                break

    # Priority: 3. Named 'TitlePlaceholder' or actual title placeholder on the first actual slide (if not in master/layout)
    if not template_title_ref_shape and template_prs.slides:
        actual_slide = template_prs.slides[0]
        if actual_slide.has_title and actual_slide.shapes.title:
            template_title_ref_shape = actual_slide.shapes.title
            logger.debug(f"Using actual title shape from template's first slide: '{template_title_ref_shape.name}'")
        else:
            for shape_on_slide in actual_slide.shapes:
                if shape_on_slide.name == "TitlePlaceholder":
                    template_title_ref_shape = shape_on_slide
                    logger.debug(f"Found named shape 'TitlePlaceholder' on template's first slide.")
                    break

    if not template_title_ref_shape or not template_title_ref_shape.has_text_frame:
        logger.warning("Could not find a suitable reference title placeholder with a text frame in the template. Applying basic text only.")
        title_shape.text = title_text
        return

    template_tf = template_title_ref_shape.text_frame
    logger.debug(f"Using template text_frame from shape: '{template_title_ref_shape.name}' (ID: {template_title_ref_shape.shape_id}) for styling.")

    # --- Copy Shape Fill --- 
    if template_title_ref_shape:
        source_shape_fill = template_title_ref_shape.fill
        target_shape_fill = title_shape.fill

        # Check if source_shape_fill.type is None first, as it can be for inherited fills
        if source_shape_fill.type is None or source_shape_fill.type == MSO_FILL_TYPE.BACKGROUND:
            target_shape_fill.background() # Makes it transparent / no fill
            logger.debug("Template Title - Source fill is NONE or BACKGROUND. Applied no fill to target.")
        elif source_shape_fill.type == MSO_FILL_TYPE.SOLID:
            target_shape_fill.solid()
            if source_shape_fill.fore_color.type == MSO_COLOR_TYPE.RGB:
                target_shape_fill.fore_color.rgb = source_shape_fill.fore_color.rgb
                logger.debug(f"Template Title - Source fill is SOLID RGB: {source_shape_fill.fore_color.rgb}. Applied to target.")
            elif source_shape_fill.fore_color.type == MSO_COLOR_TYPE.SCHEME:
                target_shape_fill.fore_color.theme_color = source_shape_fill.fore_color.theme_color
                if hasattr(source_shape_fill.fore_color, 'brightness') and source_shape_fill.fore_color.brightness is not None:
                    target_shape_fill.fore_color.brightness = source_shape_fill.fore_color.brightness
                if hasattr(source_shape_fill.fore_color, 'tint') and source_shape_fill.fore_color.tint is not None:
                    target_shape_fill.fore_color.tint = source_shape_fill.fore_color.tint
                if hasattr(source_shape_fill.fore_color, 'shade') and source_shape_fill.fore_color.shade is not None:
                    target_shape_fill.fore_color.shade = source_shape_fill.fore_color.shade
                logger.debug(f"Template Title - Source fill is SOLID SCHEME Color. Applied to target.")
            else:
                logger.warning(f"Template Title - Source fill is SOLID but color type {source_shape_fill.fore_color.type} not handled for fill. Fill color not fully copied.")
        # Add more fill types like GRADIENT, PICTURE if necessary later
        else:
            logger.warning(f"Template Title - Source fill type {source_shape_fill.type} not handled. Fill not copied.")
    else:
        logger.warning("Could not find reference title placeholder shape to copy fill.")
    # --- End of Shape Fill Copying ---

    # --- Apply Vertical Anchor from template_tf to title_shape.text_frame --- #
    source_va = template_tf.vertical_anchor
    logger.debug(f"Template Title - Read source vertical_anchor: {source_va} (Enum: {MSO_VERTICAL_ANCHOR(source_va) if source_va is not None else 'None'})")
    if source_va is not None:
        title_shape.text_frame.vertical_anchor = source_va
        logger.debug(f"Applied vertical_anchor '{source_va}' to title_shape.text_frame. New VA: {title_shape.text_frame.vertical_anchor}")
    else:
        logger.debug("Source vertical_anchor is None. Vertical alignment on new title will use default (likely TOP).")

    # --- Clear existing content and add new paragraph for the title text --- #
    title_shape.text_frame.clear()
    p = title_shape.text_frame.add_paragraph()
    p.text = title_text
    logger.debug(f"Set title text to: '{title_text}' in new paragraph.")

    # --- Apply Paragraph and Font Styling from template_tf to the new paragraph 'p' and its first run --- #
    if template_tf.paragraphs:
        ref_para = template_tf.paragraphs[0]  # Use first paragraph of template for style reference
        logger.debug(f"Reference paragraph from template: Level={ref_para.level}, SpaceBefore={ref_para.space_before}, SpaceAfter={ref_para.space_after}, LineSpacing={ref_para.line_spacing}")

        # 1. Apply Paragraph Alignment (Horizontal)
        if ref_para.alignment is not None:
            p.alignment = ref_para.alignment
            logger.debug(f"Template Title - Source Paragraph Alignment: {ref_para.alignment} (Enum: {PP_ALIGN(ref_para.alignment) if ref_para.alignment is not None else 'None'}). Applied to new paragraph.")
        else:
            logger.debug("Template Title - Source paragraph alignment is None. Horizontal alignment will use default.")
        
        # Apply other paragraph properties if needed (e.g., level, spacing - though these are often complex)
        # p.level = ref_para.level # Be cautious with level, can affect master styles

        # 2. Apply Font Styling from the first run of the reference paragraph
        if ref_para.runs:
            ref_run = ref_para.runs[0]  # Reference run from template
            
            if p.runs: # The new paragraph 'p' should have one run after p.text = title_text
                target_run = p.runs[0]

                font_attrs_log = []
                # Font Name
                if ref_run.font.name:
                    target_run.font.name = ref_run.font.name
                    font_attrs_log.append(f"Name='{ref_run.font.name}'")
                # Font Size
                if ref_run.font.size:
                    target_run.font.size = ref_run.font.size
                    font_attrs_log.append(f"Size={ref_run.font.size}")
                # Bold
                if ref_run.font.bold is not None:
                    target_run.font.bold = ref_run.font.bold
                    font_attrs_log.append(f"Bold={ref_run.font.bold}")
                # Italic
                if ref_run.font.italic is not None:
                    target_run.font.italic = ref_run.font.italic
                    font_attrs_log.append(f"Italic={ref_run.font.italic}")
                # Underline
                if ref_run.font.underline is not None:
                    target_run.font.underline = ref_run.font.underline
                    font_attrs_log.append(f"Underline={ref_run.font.underline}")
                # Color
                if ref_run.font.color.type:
                    if ref_run.font.color.type == MSO_COLOR_TYPE.RGB:
                        target_run.font.color.rgb = ref_run.font.color.rgb
                        font_attrs_log.append(f"ColorRGB='{ref_run.font.color.rgb}'")
                    elif ref_run.font.color.type == MSO_COLOR_TYPE.SCHEME:
                        target_run.font.color.theme_color = ref_run.font.color.theme_color
                        target_run.font.color.brightness = ref_run.font.color.brightness
                        font_attrs_log.append(f"ColorTheme='{ref_run.font.color.theme_color}', Brightness='{ref_run.font.color.brightness}'")
                    # Add other color types if necessary
                
                logger.debug(f"Template Title - Source Font Details: {', '.join(font_attrs_log)}")
                applied_font_attrs_log = [
                    f"Name='{target_run.font.name}'", f"Size={target_run.font.size}", 
                    f"Bold={target_run.font.bold}", f"Italic={target_run.font.italic}", 
                    f"Underline={target_run.font.underline}"
                ]
                if target_run.font.color.type == MSO_COLOR_TYPE.RGB:
                    applied_font_attrs_log.append(f"ColorRGB='{target_run.font.color.rgb}'")
                elif target_run.font.color.type == MSO_COLOR_TYPE.SCHEME:
                    applied_font_attrs_log.append(f"ColorTheme='{target_run.font.color.theme_color}', Brightness='{target_run.font.color.brightness}'")
                logger.debug(f"Applied Font to new run: {', '.join(applied_font_attrs_log)}")

            else:
                logger.warning("Title paragraph has no runs after setting text. Cannot apply font styling.")
        else:
            logger.debug("Template Title - Reference paragraph in template has no runs. Cannot copy font styling.")
    else:
        logger.debug("Template Title - Reference text_frame in template has no paragraphs. Cannot copy paragraph/font styling.")

    logger.debug(f"Finished setting title for shape '{title_shape.name}'. Text frame word_wrap: {title_shape.text_frame.word_wrap}, auto_size: {title_shape.text_frame.auto_size}")

def _add_and_style_table(slide, table_data, cell_metadata, template_slide=None):
    table_pos = get_element_position('main_table')
    if not table_pos:
        logger.error("Failed to get table position coordinates")
        return False

    table_layout = TableLayout(
        placeholder_name=TABLE_PLACEHOLDER_NAME,
        shape_name=SHAPE_NAME_TABLE,
        position=table_pos,
        row_height_header=TABLE_ROW_HEIGHT_HEADER,
        row_height_body=TABLE_ROW_HEIGHT_BODY,
        row_height_subtotal=TABLE_ROW_HEIGHT_SUBTOTAL,
        column_widths=TABLE_COLUMN_WIDTHS,
        top_override=TABLE_TOP_OVERRIDE,
        height_rule_available=TABLE_HEIGHT_RULE_AVAILABLE,
        height_rule_value=WD_ROW_HEIGHT_RULE.AT_LEAST,
    )

    return presentation_add_and_style_table(
        slide,
        table_data,
        cell_metadata,
        table_layout,
        TABLE_CELL_STYLE_CONTEXT,
        logger,
    )

def _prepare_main_table_data(df, region, masterbrand):
    """
    Prepare table data and cell metadata for the main data table.
    
    Args:
        df: DataFrame with the Excel data
        region: Region filter
        masterbrand: Masterbrand filter
        
    Returns:
        tuple: (table_data, cell_metadata) or (None, None) if no data
    """
    try:
        # Filter data for this region/masterbrand combination
        filtered_df = df[
            (df['Region'] == region) &
            (df['Global Masterbrand'] == masterbrand)
        ].copy()
        
        if filtered_df.empty:
            logger.warning(f"No data found for {region} - {masterbrand}")
            return None, None
        
        # Prepare table structure
        table_data = []
        cell_metadata = {}
        
        # Header row
        header_row = ['Campaign Name', 'Budget', 'TV GRPs', 'TOTAL REACH', 'TOTAL FREQ', '%', 'Media Type'] + \
                    ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 
                     'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'] + \
                    ['Q1', 'Q2', 'Q3', 'Q4']
        table_data.append(header_row)
        
        # Process each row of data
        row_idx = 1
        for _, row in filtered_df.iterrows():
            # Build data row
            campaign_name = str(row.get('Campaign Name', ''))
            budget = f"${row.get('Budget', 0):,.0f}" if pd.notna(row.get('Budget')) else "$0"
            tv_grps = str(row.get('TV GRPs', '')) if pd.notna(row.get('TV GRPs')) else ""
            percentage = f"{row.get('Percentage', 0):.1f}%" if pd.notna(row.get('Percentage')) else "0.0%"
            media_type = str(row.get('Media Type', ''))
            
            data_row = [campaign_name, budget, tv_grps, percentage, media_type]
            
            # Add monthly data
            months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 
                     'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
            
            for col_idx, month in enumerate(months, start=5):
                month_value = row.get(month, 0)
                if pd.notna(month_value) and month_value != 0:
                    formatted_value = f"${month_value:,.0f}"
                    data_row.append(formatted_value)
                    # Store metadata for conditional formatting - align with display
                    cell_metadata[(row_idx, col_idx)] = {
                        'value': month_value,
                        'media_type': media_type,
                        'has_data': not is_empty_formatted_value(formatted_value)
                    }
                else:
                    data_row.append("")
                    cell_metadata[(row_idx, col_idx)] = {
                        'value': 0,
                        'media_type': media_type,
                        'has_data': False
                    }
            
            # Add quarterly data
            quarters = ['Q1', 'Q2', 'Q3', 'Q4']
            for col_idx, quarter in enumerate(quarters, start=17):
                quarter_value = row.get(quarter, 0)
                if pd.notna(quarter_value) and quarter_value != 0:
                    data_row.append(f"${quarter_value:,.0f}")
                else:
                    data_row.append("–")
            
            table_data.append(data_row)
            row_idx += 1
        
        # Add totals row
        total_budget = filtered_df['Budget'].sum() if 'Budget' in filtered_df.columns else 0
        totals_row = [
            'TOTAL',
            f"${total_budget:,.0f}",
            '',
            '',
            ''  # Media type blank for total row
        ] + \
        [f"${month_total:,.0f}" if month_total > 0 else "" for month_total in [filtered_df[month].sum() for month in months]] + \
        [f"${quarter_total:,.0f}" if quarter_total > 0 else "–" for quarter_total in [filtered_df[quarter].sum() for quarter in quarters]]
        
        table_data.append(totals_row)
        
        logger.info(f"Prepared table data with {len(table_data)} rows for {region} - {masterbrand}")
        return table_data, cell_metadata
        
    except Exception as e:
        logger.error(f"Error preparing table data: {str(e)}")
        logger.error(traceback.format_exc())
        return None, None

def _prepare_funnel_chart_data(df, region, masterbrand):
    """
    Prepare data for the funnel chart.
    
    Args:
        df: DataFrame with the Excel data
        region: Region filter
        masterbrand: Masterbrand filter
        
    Returns:
        list: Chart data or None if no data
    """
    try:
        # Filter data for this region/masterbrand combination
        filtered_df = df[
            (df['Region'] == region) &
            (df['Global Masterbrand'] == masterbrand)
        ].copy()
        
        if filtered_df.empty:
            return None
        
        # Group by funnel stage and sum budgets
        if 'Funnel Stage' in filtered_df.columns and 'Budget' in filtered_df.columns:
            funnel_data = filtered_df.groupby('Funnel Stage')['Budget'].sum().to_dict()
            return [(stage, budget) for stage, budget in funnel_data.items() if budget > 0]
        
        return None
        
    except Exception as e:
        logger.error(f"Error preparing funnel chart data: {str(e)}")
        return None

def _prepare_media_type_chart_data(df, region, masterbrand):
    """
    Prepare data for the media type chart.
    
    Args:
        df: DataFrame with the Excel data
        region: Region filter
        masterbrand: Masterbrand filter
        
    Returns:
        list: Chart data or None if no data
    """
    try:
        # Filter data for this region/masterbrand combination
        filtered_df = df[
            (df['Region'] == region) &
            (df['Global Masterbrand'] == masterbrand)
        ].copy()
        
        if filtered_df.empty:
            return None
        
        # Group by media type and sum budgets
        if 'Media Type' in filtered_df.columns and 'Budget' in filtered_df.columns:
            media_data = filtered_df.groupby('Media Type')['Budget'].sum().to_dict()
            return [(media_type, budget) for media_type, budget in media_data.items() if budget > 0]
        
        return None
        
    except Exception as e:
        logger.error(f"Error preparing media type chart data: {str(e)}")
        return None

def _prepare_brand_chart_data(df, region, masterbrand):
    """
    Prepare data for the brand chart.
    
    Args:
        df: DataFrame with the Excel data
        region: Region filter
        masterbrand: Masterbrand filter
        
    Returns:
        list: Chart data or None if no data
    """
    try:
        # Filter data for this region/masterbrand combination
        filtered_df = df[
            (df['Region'] == region) &
            (df['Global Masterbrand'] == masterbrand)
        ].copy()
        
        if filtered_df.empty:
            return None
        
        # Group by brand and sum budgets
        if 'Brand' in filtered_df.columns and 'Budget' in filtered_df.columns:
            brand_data = filtered_df.groupby('Brand')['Budget'].sum().to_dict()
            return [(brand, budget) for brand, budget in brand_data.items() if budget > 0]
        
        return None
        
    except Exception as e:
        logger.error(f"Error preparing brand chart data: {str(e)}")
        return None

def _add_pie_chart(slide, chart_data, chart_title, position_info, chart_name=None):
    """Compatibility wrapper for modular pie chart generation."""

    if not position_info:
        logger.error("Failed to get chart position coordinates")
        return False

    return presentation_add_pie_chart(
        slide,
        chart_data,
        chart_title,
        position_info,
        CHART_STYLE_CONTEXT,
        CHART_COLOR_MAPPING,
        CHART_COLOR_CYCLE,
        chart_name=chart_name,
    )

def _populate_slide_content(new_slide, prs, combination_row, slide_title_suffix, 
                          split_table_data, split_metadata, split_idx, df, excel_path):
    """Populate a single slide with all content (title, table, charts, comments)."""
    
    # Copy Title Placeholder from template FIRST
    template_title_shape = _get_shape_by_name(prs.slides[0], SHAPE_NAME_TITLE)
    if template_title_shape:
        title_text = f"{combination_row[0]} – {combination_row[1]} ({combination_row[2]}){slide_title_suffix}"
        copied_title_shape = _copy_text_box(template_title_shape, new_slide, new_name=SHAPE_NAME_TITLE, new_text=title_text)
        if copied_title_shape:
            # QA FIX: Ensure correct font for title
            if copied_title_shape.has_text_frame and copied_title_shape.text_frame.paragraphs:
                for paragraph in copied_title_shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        # PRIORITY 3: Enhanced font consistency for titles
                        _ensure_font_consistency(
                            run.font,
                            target_font_name=DEFAULT_FONT_NAME,
                            target_size=FONT_SIZE_TITLE,
                            target_bold=True,
                            target_color=CLR_WHITE
                        )
            logger.info(f"Title copied and populated for slide.")
    
    # Copy Static Text Elements (Comments Title and Box)
    template_comments_title_shape = _get_shape_by_name(prs.slides[0], SHAPE_NAME_COMMENTS_TITLE)
    if template_comments_title_shape:
        copied_comments_title = _copy_text_box(template_comments_title_shape, new_slide, new_name=SHAPE_NAME_COMMENTS_TITLE, new_text="COMMENTS")
        if copied_comments_title:
            # Position comments title using centralized coordinates
            title_pos = get_element_position('comments_title')
            if title_pos:
                copied_comments_title.left = title_pos['left']
                copied_comments_title.top = title_pos['top']
                copied_comments_title.width = title_pos['width']
                copied_comments_title.height = title_pos['height']
    
    # Copy Comments Box
    template_comments_box_shape = _get_shape_by_name(prs.slides[0], SHAPE_NAME_COMMENTS_BOX)
    if template_comments_box_shape:
        copied_comments_box = _copy_text_box(template_comments_box_shape, new_slide, new_name=SHAPE_NAME_COMMENTS_BOX, new_text="")
        if copied_comments_box:
            # Position comments box using centralized coordinates
            box_pos = get_element_position('comments_box')
            if box_pos:
                copied_comments_box.left = box_pos['left']
                copied_comments_box.top = box_pos['top']
                copied_comments_box.width = box_pos['width']
                copied_comments_box.height = box_pos['height']
    
    logger.info("Skipping legend copying - legend elements are part of slide master")
    
    # Create and populate the main data table
    logger.info(f"Creating table for {combination_row[0]} - {combination_row[1]} - {combination_row[2]}{slide_title_suffix}")
    
    table_success = _add_and_style_table(new_slide, split_table_data, split_metadata, prs.slides[0])
    if table_success:
        logger.info(f"Table created successfully for slide")
    else:
        logger.warning(f"Failed to create table for slide")
    
    # Create and add the three pie charts
    if SHOW_CHARTS_ON_SPLITS == "all" or (split_idx == 0 and SHOW_CHARTS_ON_SPLITS == "first_only"):
        logger.info(f"Creating charts for {combination_row[0]} - {combination_row[1]} - {combination_row[2]}")
        
        # Chart positioning using centralized coordinate system
        chart_positions = [
            get_element_position('chart_1'),
            get_element_position('chart_2'),
            get_element_position('chart_3')
        ]
        
        # Validate chart positions
        if not all(chart_positions):
            logger.error("Failed to get chart position coordinates")
            chart_positions = []
        
        # 1. Funnel Chart
        funnel_data = _prepare_funnel_chart_data_detailed(df, combination_row[0], combination_row[1], combination_row[2])
        if funnel_data:
            funnel_success = _add_pie_chart(new_slide, funnel_data, "BUDGET BY FUNNEL STAGE", chart_positions[0], SHAPE_NAME_FUNNEL_CHART)
            if funnel_success:
                logger.info(f"Funnel chart created successfully for slide")
        
        # 2. Media Type Chart  
        media_type_data = _prepare_media_type_chart_data_detailed(df, combination_row[0], combination_row[1], combination_row[2])
        if media_type_data:
            media_success = _add_pie_chart(new_slide, media_type_data, "BUDGET BY MEDIA TYPE", chart_positions[1], SHAPE_NAME_MEDIA_TYPE_CHART)
            if media_success:
                logger.info(f"Media type chart created successfully for slide")
        
        # 3. Campaign Type Chart
        campaign_type_data = _prepare_campaign_type_chart_data(df, combination_row[0], combination_row[1], combination_row[2])
        if campaign_type_data:
            campaign_success = _add_pie_chart(new_slide, campaign_type_data, "BUDGET BY CAMPAIGN TYPE", chart_positions[2], SHAPE_NAME_CAMPAIGN_TYPE_CHART)
            if campaign_success:
                logger.info(f"Campaign type chart created successfully for slide")
                
        logger.info(f"Chart creation completed for slide")

def create_presentation(template_path, excel_path, output_path):
    """Creates a PowerPoint presentation based on a template and Excel data."""
    logger.info(f"Starting presentation creation using template: {template_path}")
    try:
        prs = Presentation(template_path)
        if not prs.slides:
            logger.error("Template presentation has no slides. Cannot proceed.")
            return False

        logger.debug("--- Available Slide Layouts in Template ---")
        for i, layout in enumerate(prs.slide_layouts):
            layout_name = layout.name if hasattr(layout, 'name') else 'Unknown Name'
            placeholder_count = len(layout.placeholders) if hasattr(layout, 'placeholders') else 0
            logger.debug(f"  Layout Index {i}: {layout_name} (Placeholders: {placeholder_count})")
        logger.debug("--- End Available Slide Layouts ---")

        # Log template slide structure
        logger.debug("--- Template Slide (First Slide) Structure ---")
        if prs.slides:
            template_slide = prs.slides[0]
            logger.debug(f"  Template slide has {len(template_slide.shapes)} shapes:")
            for i, shape in enumerate(template_slide.shapes):
                shape_name = shape.name if hasattr(shape, 'name') else 'Unnamed'
                shape_type = shape.shape_type if hasattr(shape, 'shape_type') else 'Unknown Type'
                logger.debug(f"    Shape {i}: Name='{shape_name}', Type={shape_type}")
        else:
            logger.debug("  No slides found in template!")
        logger.debug("--- End Template Slide Structure ---")

        df = load_and_prepare_data(excel_path)
        if df is None or df.empty:
            logger.error("Failed to load or prepare data. Aborting presentation creation.")
            return False

        # The load_and_prepare_data() function returns a processed DataFrame with these column names
        country_col_name = 'Country'
        brand_col_name = 'Brand'
        year_col_name = 'Year'
        
        unique_combinations = df[[country_col_name, brand_col_name, year_col_name]].drop_duplicates().values.tolist()
        logger.info(f"Found {len(unique_combinations)} unique Country/Global Masterbrand/Year combinations.")

        # Calculate total investment for each combination
        combinations_with_investment = []
        for combination in unique_combinations:
            country, brand, year = combination
            # Filter data for this specific combination
            combination_data = df[
                (df[country_col_name] == country) & 
                (df[brand_col_name] == brand) & 
                (df[year_col_name] == year)
            ]
            # Calculate total investment (sum of Total Cost for this combination)
            total_investment = combination_data['Total Cost'].fillna(0).sum()
            
            combinations_with_investment.append((country, brand, year, total_investment))
        
        # Group by market (country) and calculate total market investment
        market_investments = {}
        for country, brand, year, investment in combinations_with_investment:
            if country not in market_investments:
                market_investments[country] = {'total': 0, 'combinations': []}
            market_investments[country]['total'] += investment
            market_investments[country]['combinations'].append((country, brand, year, investment))
        
        # Sort markets by total investment (highest first)
        sorted_markets = sorted(market_investments.items(), key=lambda x: x[1]['total'], reverse=True)
        
        # Build final sorted list: markets ordered by total, brands within market ordered by individual investment
        unique_combinations = []
        for market, data in sorted_markets:
            # Sort combinations within this market by individual investment
            market_combos = sorted(data['combinations'], key=lambda x: x[3], reverse=True)
            # Add to final list (without investment values)
            unique_combinations.extend([(c, b, y) for c, b, y, _ in market_combos])
        
        # Log the sorted order with market totals
        logger.info("Slides will be generated grouped by market, ordered by total market investment:")
        for i, (market, data) in enumerate(sorted_markets[:10]):  # Log top 10 markets
            logger.info(f"  {i+1}. {market}: £{data['total']:,.0f} total")
            for combo in sorted(data['combinations'], key=lambda x: x[3], reverse=True)[:3]:  # Show top 3 brands
                logger.info(f"      - {combo[1]} ({combo[2]}): £{combo[3]:,.0f}")
            if len(data['combinations']) > 3:
                logger.info(f"      ... and {len(data['combinations']) - 3} more brands")
        if len(sorted_markets) > 10:
            logger.info(f"  ... and {len(sorted_markets) - 10} more markets")

        if not unique_combinations:
            logger.warning("No unique Country/Global Masterbrand combinations found in the data.")
            # Decide if an empty presentation should be saved or an error returned
            # For now, let's save an empty presentation (after removing template slide)

        current_market = None
        for idx, combination_row in enumerate(unique_combinations):
            # Check if we're starting a new market
            if combination_row[0] != current_market:
                current_market = combination_row[0]
                # Fix market name display
                display_market_name = "Morocco" if current_market == "MOR" else current_market
                # Add a market delimiter slide with black background
                try:
                    # Try to use a blank slide layout 
                    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
                    delimiter_slide = prs.slides.add_slide(blank_layout)
                except:
                    delimiter_slide = prs.slides.add_slide(prs.slide_layouts[0])
                
                # Add full black background
                slide_width = prs.slide_width
                slide_height = prs.slide_height
                
                # Create a black rectangle covering the entire slide
                from pptx.enum.shapes import MSO_SHAPE
                black_bg = delimiter_slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    left=0,
                    top=0,
                    width=slide_width,
                    height=slide_height
                )
                
                # Make the background solid black
                black_bg.fill.solid()
                black_bg.fill.fore_color.rgb = RGBColor(0, 0, 0)  # Pure black
                black_bg.line.fill.background()  # Remove border
                
                # Create a text box in the center of the slide for the market name
                text_box = delimiter_slide.shapes.add_textbox(
                    left=Inches(1),
                    top=int(slide_height * 0.45),  # Center vertically
                    width=slide_width - Inches(2),
                    height=Inches(1.5)
                )
                
                text_frame = text_box.text_frame
                text_frame.text = str(display_market_name).upper() if display_market_name else "UNKNOWN"  # Make market name full caps
                text_frame.word_wrap = False
                text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                
                # Format the text - large white font
                for paragraph in text_frame.paragraphs:
                    paragraph.alignment = PP_ALIGN.CENTER
                    for run in paragraph.runs:
                        run.font.size = Pt(36)  # Large font (30-40 range)
                        run.font.bold = True
                        run.font.name = DEFAULT_FONT_NAME
                        run.font.color.rgb = RGBColor(255, 255, 255)  # Pure white
                logger.info(f"Added market delimiter slide for: {display_market_name}")
            
            logger.info(f"Processing combination {idx+1}/{len(unique_combinations)}: {combination_row[0]} - {combination_row[1]} - {combination_row[2]}")
            
            # First, prepare the table data to check if splitting is needed
            table_result = _prepare_main_table_data_detailed(df, combination_row[0], combination_row[1], combination_row[2], excel_path)
            
            if table_result[0] is None:
                logger.warning(f"No table data generated for {combination_row[0]} - {combination_row[1]}")
                continue
            
            table_data, cell_metadata = table_result
            
            # Split table if needed
            table_splits = _split_table_data_by_campaigns(table_data, cell_metadata)
            
            # Create a slide for each split
            for split_idx, (split_table_data, split_metadata, is_continuation) in enumerate(table_splits):
                # Add slide number to title if there are multiple splits
                if len(table_splits) > 1:
                    slide_title_suffix = f" ({split_idx + 1} of {len(table_splits)})"
                else:
                    slide_title_suffix = ""
                
                new_slide = prs.slides.add_slide(prs.slide_layouts[0])
                # logger.debug(f"Added new slide for {combination_row[0]} - {combination_row[1]} - {combination_row[2]}{slide_title_suffix}")
                
                # Populate this slide with content immediately
                _populate_slide_content(
                    new_slide, prs, combination_row, slide_title_suffix, 
                    split_table_data, split_metadata, split_idx, df, excel_path
                )

                # Diagnostic: Log all shapes on the new slide (commented out to reduce log size)
            # This entire block is commented out to reduce log file size
            # Re-enable for debugging shape-related issues
            """
            logger.debug(f"--- Shapes on new_slide (Layout: {prs.slide_layouts[0].name}, Index: {idx+1}) ---")
            if hasattr(new_slide, 'shapes'):
                logger.debug(f"  new_slide.shapes attribute exists. Number of shapes found: {len(new_slide.shapes)}")
                if new_slide.shapes:
                    for i, shape in enumerate(new_slide.shapes):
                        shape_name = shape.name if hasattr(shape, 'name') else 'Unnamed Shape'
                        shape_type = shape.shape_type if hasattr(shape, 'shape_type') else 'Unknown Type'
                        is_placeholder = shape.is_placeholder if hasattr(shape, 'is_placeholder') else False
                        placeholder_type = None
                        placeholder_idx = None
                        if is_placeholder and hasattr(shape, 'placeholder_format'):
                            ph_format = shape.placeholder_format
                            if hasattr(ph_format, 'type'):
                                placeholder_type = ph_format.type
                            if hasattr(ph_format, 'idx'):
                                placeholder_idx = ph_format.idx
                        logger.debug(f"  Shape {i}: Name='{shape_name}', Type={shape_type}, IsPlaceholder={is_placeholder}, PlaceholderType={placeholder_type}, PlaceholderIdx={placeholder_idx}")
                else:
                    logger.debug("  new_slide.shapes collection is empty (it exists but len is 0).")
            else:
                logger.debug("  new_slide.shapes attribute is missing.")
            logger.debug(f"--- End Shapes on new_slide (Index: {idx+1}) ---")
                """

        # ORIGINAL CONTENT POPULATION CODE MOVED TO _populate_slide_content() FUNCTION
        # (This section was deleted to prevent empty slides)
        # Remove the original template slide (slide 1) which shows "Region / Global Masterbrand"
        if len(prs.slides) > 1: # Only remove if we have other slides
            logger.info(f"Removing template slide from presentation (total slides: {len(prs.slides)}).")
            try:
                # The template slide is always the first slide added (index 0)
                rId = prs.slides._sldIdLst[0].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[0]
                logger.info("Template slide removed successfully.")
            except Exception as e:
                logger.warning(f"Could not remove template slide: {e}")
        else:
            logger.warning(f"Not removing template slide - only {len(prs.slides)} slides found.")

        # CRITICAL FIX: Ensure output directory exists and add proper save error handling
        from pathlib import Path
        output_path_obj = Path(output_path)
        
        # Ensure output directory exists
        output_path_obj.parent.mkdir(parents=True, exist_ok=True)
        logger.info(f"Output directory ensured: {output_path_obj.parent}")
        
        # Ensure .pptx extension
        if not output_path.lower().endswith('.pptx'):
            output_path = output_path + '.pptx'
            logger.info(f"Added .pptx extension: {output_path}")

        # Save with proper error handling
        try:
            prs.save(output_path)
            logger.info(f"Presentation saved to {output_path}")
            
            # Verify file creation and get size
            file_size = os.path.getsize(output_path)
            logger.info(f"File verified: {file_size:,} bytes")
            return True
        except Exception as e:
            logger.exception(f"❌ Save failed: {e}")
            return False

        return True

    except Exception as e:
        logger.error(f"An error occurred during presentation creation: {e}")
        logger.error(traceback.format_exc())
        return False

def _apply_internal_table_borders(table, total_rows):
    """
    Apply specific internal borders for the first 7 columns only (excluding header and total rows).
    Uses direct OXML manipulation to add #BFBFBF borders with 0.75pt thickness.
    
    Args:
        table: The PowerPoint table object
        total_rows: Total number of rows in the table
    """
    try:
        from pptx.oxml.ns import qn
        from pptx.oxml import parse_xml
        from pptx.dml.color import RGBColor
        from pptx.util import Pt
        
        # Border specifications - same as external borders for consistency
        border_color_rgb = (191, 191, 191) # CLR_TABLE_GRAY is RGBColor(191, 191, 191)
        border_width_emu = int(0.75 * 12700)  # 0.75pt in EMUs
        hex_color = f"{border_color_rgb[0]:02X}{border_color_rgb[1]:02X}{border_color_rgb[2]:02X}"
        
        logger.debug(f"Applying internal borders to first 7 columns for rows 1 to {total_rows-2}")
        
        # Apply internal borders only to body rows (excluding header row 0 and total row)
        for row_idx in range(1, total_rows - 1):  # Skip header (0) and total (last) rows
            row = table.rows[row_idx]
            
            for col_idx in range(7):  # First 7 columns only
                cell = row.cells[col_idx]
                
                try:
                    # Access the cell's table cell properties
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    
                    # Add bottom border for horizontal lines (apply to all qualifying cells)
                    if row_idx < total_rows - 2:  # Don't add bottom border to row just above total
                        bottom_border = tcPr.find(qn('a:lnB'))
                        if bottom_border is None:
                            bottom_border = tcPr.add(qn('a:lnB'))
                        
                        # Set border properties
                        bottom_border.set('w', str(border_width_emu))
                        bottom_border.set('cap', 'flat')
                        bottom_border.set('cmpd', 'sng')
                        
                        # Add solid fill with color
                        solidFill = bottom_border.find(qn('a:solidFill'))
                        if solidFill is None:
                            solidFill = bottom_border.add(qn('a:solidFill'))
                        
                        srgbClr = solidFill.find(qn('a:srgbClr'))
                        if srgbClr is None:
                            srgbClr = solidFill.add(qn('a:srgbClr'))
                        srgbClr.set('val', hex_color)
                        
                        # Add preset dash (solid line)
                        prstDash = bottom_border.find(qn('a:prstDash'))
                        if prstDash is None:
                            prstDash = bottom_border.add(qn('a:prstDash'))
                        prstDash.set('val', 'solid')
                    
                    # Add right border for vertical lines (apply to first 6 columns only)
                    if col_idx < 6:  # Columns 0, 1, 2, 3, 4, 5 get right borders
                        right_border = tcPr.find(qn('a:lnR'))
                        if right_border is None:
                            right_border = tcPr.add(qn('a:lnR'))
                        
                        # Set border properties
                        right_border.set('w', str(border_width_emu))
                        right_border.set('cap', 'flat')
                        right_border.set('cmpd', 'sng')
                        
                        # Add solid fill with color
                        solidFill = right_border.find(qn('a:solidFill'))
                        if solidFill is None:
                            solidFill = right_border.add(qn('a:solidFill'))
                        
                        srgbClr = solidFill.find(qn('a:srgbClr'))
                        if srgbClr is None:
                            srgbClr = solidFill.add(qn('a:srgbClr'))
                        srgbClr.set('val', hex_color)
                        
                        # Add preset dash (solid line)
                        prstDash = right_border.find(qn('a:prstDash'))
                        if prstDash is None:
                            prstDash = right_border.add(qn('a:prstDash'))
                        prstDash.set('val', 'solid')
                    
                except Exception as cell_border_error:
                    logger.debug(f"Could not apply internal borders to cell ({row_idx}, {col_idx}): {cell_border_error}")
        
        logger.info(f"Internal table borders applied successfully for first 7 columns with #BFBFBF color and 0.75pt width")
        return True
        
    except Exception as e:
        logger.warning(f"Error applying internal table borders: {str(e)}")
        return False

def _verify_file_exists(label, path_str, extra_search=()):
    """
    Return an absolute path to an existing file.

    extra_search – optional directories (relative to project root) to probe
                   if the user gave a bare filename.
    """
    from pathlib import Path
    p = Path(path_str).expanduser()
    
    # Try the path as given first
    if p.is_file():
        return str(p.resolve())
    
    # If the file doesn't exist at the given path and we have fallback directories,
    # and if it looks like a bare filename (no directory separators), search fallbacks
    if extra_search and '/' not in str(p) and '\\' not in str(p):
        project_root = Path(__file__).parent.parent  # Go up from scripts/ to project root
        for alt in extra_search:
            candidate = project_root / alt / p.name
            if candidate.is_file():
                return str(candidate.resolve())

    # If we get here, the file wasn't found anywhere
    p_resolved = p.resolve()
    raise FileNotFoundError(f"❌ {label} file not found: '{p_resolved}'")

def _unit_test__no_orphan_self():
    import inspect, re, pathlib
    src = pathlib.Path(__file__).read_text(encoding="utf-8")
    tree = ast.parse(src, filename=str(pathlib.Path(__file__)))
    offender_lines = []
    for node in ast.walk(tree):
        # Set parent attributes for all nodes to allow easy traversal up the tree
        for child in ast.iter_child_nodes(node):
            child.parent = node
            
        if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name) and node.value.id == 'self':
            # Ascend until we hit a FunctionDef, ClassDef, or Module
            current_node = node
            legit = False
            is_in_class_scope = False # Track if we are within a ClassDef scope at all
            
            while hasattr(current_node, 'parent'): # Check if parent attribute exists
                parent = current_node.parent
                if isinstance(parent, ast.FunctionDef):
                    # Check if 'self' is the first argument of this function
                    if parent.args.args and parent.args.args[0].arg == 'self':
                        # Now, ensure this FunctionDef is itself within a ClassDef
                        # Keep ascending from the FunctionDef to find a ClassDef
                        func_parent = parent
                        while hasattr(func_parent, 'parent'):
                            func_parent = func_parent.parent
                            if isinstance(func_parent, ast.ClassDef):
                                legit = True # 'self' is in a method of a class
                                break
                            if isinstance(func_parent, ast.Module): # Reached top level without finding a class
                                break # Stop ascent once a FunctionDef is found and checked
                
                if isinstance(parent, ast.ClassDef):
                    is_in_class_scope = True # We are inside a class, but 'self' might be at class level (e.g. self.x = 10)
                    # If 'self' is used directly under ClassDef (not in a method), it's an error unless it's part of an assignment in __init__ or other method
                    # The current check correctly flags 'self.attribute' if not in a method like 'def meth(self, ...):'
                    # If we hit a ClassDef before a qualifying FunctionDef that has 'self' as first arg, it's likely a class variable access using 'self', which is wrong.
                    break # Stop ascent if ClassDef is found before a qualifying FunctionDef
                
                if isinstance(parent, ast.Module):
                    break # Reached top level
                current_node = parent
            
            if not legit:
                # Additional check: if it's in a class scope but not in a method, it's an error
                # e.g. class MyClass: self.x = 10 (invalid)
                # vs class MyClass: def __init__(self): self.x = 10 (valid)
                # The 'legit' flag already covers the valid case. If not legit, it's an offender.
                offender_lines.append(node.lineno)

    if offender_lines:
        # Remove duplicates and sort
        unique_offender_lines = sorted(list(set(offender_lines)))
        raise RuntimeError(f" orphan 'self.' detected outside a method or in an invalid class context at lines: {unique_offender_lines} – aborting build")

_unit_test__no_orphan_self()


def build_presentation(template_path, excel_path, output_path):
    """Backward-compatible wrapper around ``create_presentation``."""

    return create_presentation(template_path, excel_path, output_path)
