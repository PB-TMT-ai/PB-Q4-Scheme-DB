"""
Standalone Slab Report Generator
================================
Generates a formatted Excel report with slab analysis based on
qualified points. Outputs to .workspace/ directory.

Usage:
    python scripts/generate_slab_report.py
"""

from __future__ import annotations

import glob
import os
import sys
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Add project root to path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.lib.logger import error as log_error
from src.lib.logger import info as log_info

# ---------------------------------------------------------------------------
# CONSTANTS (mirrored from app.py)
# ---------------------------------------------------------------------------

DATA_DIR: str = str(PROJECT_ROOT / "data")
OUTPUT_DIR: str = str(PROJECT_ROOT / ".workspace")
SHEET_NAME: str = "Sheet1"
HEADER_ROW: int = 0

MONTH_LABELS: list[str] = [
    "Apr", "May", "Jun", "Jul", "Aug", "Sep",
    "Oct", "Nov", "Dec", "Jan", "Feb", "Mar",
]

SLAB_CONFIG: list[dict] = [
    {"slab": "Unqualified", "range": "0 – 749", "lower": 0, "upper": 750,
     "gift": "No Gift", "gift_full": "Below minimum qualification",
     "category": "Unqualified", "color": "#94a3b8"},
    {"slab": "Slab A", "range": "750 – 2,999", "lower": 750, "upper": 3000,
     "gift": "Foot Massager", "gift_full": "Foot massager",
     "category": "A", "color": "#f59e0b"},
    {"slab": "Slab B", "range": "3,000 – 4,199", "lower": 3000, "upper": 4200,
     "gift": "Sony Sound Bar", "gift_full": "Sony - sound bar, woofer and speakers",
     "category": "B", "color": "#6366f1"},
    {"slab": "Slab C", "range": "4,200 – 6,799", "lower": 4200, "upper": 6800,
     "gift": "Robot Vacuum", "gift_full": "Robot Vacuum cleaner",
     "category": "C", "color": "#10b981"},
    {"slab": "Slab D", "range": "6,800 – 7,499", "lower": 6800, "upper": 7500,
     "gift": "Apple iPad", "gift_full": "Apple iPad",
     "category": "D", "color": "#3b82f6"},
    {"slab": "Slab E", "range": "7,500+", "lower": 7500, "upper": float("inf"),
     "gift": "Washing Machine", "gift_full": "Samsung front-load washing machine",
     "category": "E", "color": "#ec4899"},
]

SLAB_ORDER: list[str] = [s["slab"] for s in SLAB_CONFIG]

POINTS_CONFIG: dict = {
    "min_volume_mt": 30,
    "points_per_mt": 25,
    "milestones": [
        {"threshold": 200, "bonus_pct": 50},
        {"threshold": 150, "bonus_pct": 30},
        {"threshold": 100, "bonus_pct": 20},
        {"threshold": 50, "bonus_pct": 10},
    ],
}

# ---------------------------------------------------------------------------
# HELPERS (mirrored from app.py)
# ---------------------------------------------------------------------------


def calculate_qualified_volume(shop_vol: float, site_vol: float) -> float:
    """Calculate qualified volume from shop and site volumes.

    Args:
        shop_vol: Total shop volume.
        site_vol: Total site volume (pre-capped).

    Returns:
        Combined qualified volume.
    """
    shop = float(shop_vol) if not pd.isna(shop_vol) else 0.0
    site = float(site_vol) if not pd.isna(site_vol) else 0.0
    return shop + site


def calculate_qualified_points(qualified_vol_mt: float) -> float:
    """Calculate qualified points from qualified volume.

    Args:
        qualified_vol_mt: Qualified volume in MT.

    Returns:
        Calculated points (rounded to nearest integer).
    """
    if pd.isna(qualified_vol_mt):
        qualified_vol_mt = 0.0
    vol = float(qualified_vol_mt)

    if vol < POINTS_CONFIG["min_volume_mt"]:
        return 0.0

    base_points = vol * POINTS_CONFIG["points_per_mt"]

    bonus_pct = 0
    for milestone in POINTS_CONFIG["milestones"]:
        if vol >= milestone["threshold"]:
            bonus_pct = milestone["bonus_pct"]
            break

    total = base_points * (1 + bonus_pct / 100)
    return round(total)


def assign_slab(points: float) -> str:
    """Assign a slab label based on qualified points.

    Args:
        points: Qualified points value.

    Returns:
        Slab label string.
    """
    if pd.isna(points):
        points = 0.0
    points = float(points)
    for cfg in SLAB_CONFIG:
        if cfg["lower"] <= points < cfg["upper"]:
            return cfg["slab"]
    return SLAB_CONFIG[0]["slab"]


def format_indian(number: float, prefix: str = "", decimal: int = 0) -> str:
    """Format a number with Indian comma grouping.

    Args:
        number: The numeric value to format.
        prefix: Optional prefix string.
        decimal: Number of decimal places.

    Returns:
        Formatted string with Indian-style commas.
    """
    if pd.isna(number):
        return f"{prefix}0"
    number = round(float(number), decimal)
    is_negative = number < 0
    number = abs(number)

    if decimal > 0:
        int_part, dec_part = f"{number:.{decimal}f}".split(".")
    else:
        int_part = str(int(number))
        dec_part = ""

    if len(int_part) <= 3:
        formatted = int_part
    else:
        last3 = int_part[-3:]
        remaining = int_part[:-3]
        groups = []
        while remaining:
            groups.append(remaining[-2:])
            remaining = remaining[:-2]
        groups.reverse()
        formatted = ",".join(groups) + "," + last3

    result = f"{prefix}{formatted}"
    if dec_part:
        result += f".{dec_part}"
    if is_negative:
        result = f"-{result}"
    return result


def _find_excel_file() -> Path:
    """Locate the most recently modified Excel file in DATA_DIR.

    Returns:
        Path to the latest .xlsx file.

    Raises:
        FileNotFoundError: If no Excel file is found.
    """
    pattern = str(Path(DATA_DIR) / "*.xlsx")
    files = glob.glob(pattern)
    if not files:
        raise FileNotFoundError(f"No .xlsx files found in {DATA_DIR}/")
    return Path(max(files, key=os.path.getmtime))


# ---------------------------------------------------------------------------
# DATA LOADING
# ---------------------------------------------------------------------------


def load_data() -> tuple[pd.DataFrame, list[str]]:
    """Load and clean Excel data (no Streamlit dependency).

    Returns:
        Tuple of (cleaned DataFrame, list of month column names).
    """
    file_path = _find_excel_file()
    log_info(f"Loading data from {file_path}")

    df = pd.read_excel(
        file_path,
        sheet_name=SHEET_NAME,
        header=HEADER_ROW,
        engine="openpyxl",
    )

    # Rename month columns
    col_map: dict[str, str] = {}
    month_cols: list[str] = []

    for col in df.columns:
        matched = False
        if hasattr(col, "strftime"):
            short = col.strftime("%b")
            if short in MONTH_LABELS and short not in col_map.values():
                col_map[col] = short
                month_cols.append(short)
                matched = True
        if not matched and isinstance(col, str):
            col_str = col.strip()
            for label in MONTH_LABELS:
                if col_str.lower().startswith(label.lower()) and label not in col_map.values():
                    col_map[col] = label
                    month_cols.append(label)
                    break

    if col_map:
        df = df.rename(columns=col_map)

    if not month_cols:
        for label in MONTH_LABELS:
            if label in df.columns:
                month_cols.append(label)

    # Coerce numerics
    for col in month_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Total Volume
    if month_cols and "Total Volume" not in df.columns:
        df["Total Volume"] = df[month_cols].sum(axis=1)

    # Shop / Site volumes
    for vol_col in ["Shop Volume", "Site Volume"]:
        if vol_col in df.columns:
            df[vol_col] = pd.to_numeric(df[vol_col], errors="coerce").fillna(0.0)
        else:
            df[vol_col] = 0.0

    # Derived columns
    df["Qualified Volume"] = df.apply(
        lambda row: calculate_qualified_volume(row["Shop Volume"], row["Site Volume"]),
        axis=1,
    )
    df["Qualified Points"] = df["Qualified Volume"].apply(calculate_qualified_points)
    df["Qualified Slab"] = df["Qualified Points"].apply(assign_slab)

    log_info(f"Loaded {len(df)} rows")
    return df, month_cols


# ---------------------------------------------------------------------------
# EXCEL REPORT GENERATION
# ---------------------------------------------------------------------------

HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def _auto_width(ws: object) -> None:
    """Auto-adjust column widths based on content.

    Args:
        ws: openpyxl worksheet object.
    """
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 3, 40)


def _style_header_row(ws: object, num_cols: int) -> None:
    """Apply header styling to the first row.

    Args:
        ws: openpyxl worksheet object.
        num_cols: Number of columns to style.
    """
    for col_idx in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER


def generate_report(df: pd.DataFrame, month_cols: list[str]) -> Path:
    """Generate a formatted Excel slab report.

    Args:
        df: Cleaned DataFrame with slab assignments.
        month_cols: List of month column names.

    Returns:
        Path to the generated report file.
    """
    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "slab_report.xlsx"

    wb = Workbook()

    # --- Sheet 1: Slab Summary ---
    ws_summary = wb.active
    ws_summary.title = "Slab Summary"
    ws_summary.sheet_properties.tabColor = "1F4E79"

    headers = ["Slab", "Points Range", "Dealer Count", "Total Volume",
               "Qualified Volume", "Total Points", "Gift"]
    ws_summary.append(headers)

    for cfg in SLAB_CONFIG:
        slab_df = df[df["Qualified Slab"] == cfg["slab"]] if "Qualified Slab" in df.columns else pd.DataFrame()
        count = len(slab_df)
        vol = slab_df["Total Volume"].sum() if "Total Volume" in slab_df.columns else 0
        qual = slab_df["Qualified Volume"].sum() if "Qualified Volume" in slab_df.columns else 0
        pts = slab_df["Qualified Points"].sum() if "Qualified Points" in slab_df.columns else 0
        ws_summary.append([
            cfg["slab"], cfg["range"], count, round(vol, 1),
            round(qual, 1), round(pts), cfg["gift_full"],
        ])

    _style_header_row(ws_summary, len(headers))
    _auto_width(ws_summary)
    ws_summary.freeze_panes = "A2"

    # --- Sheet 2: Dealer Detail ---
    ws_detail = wb.create_sheet("Dealer Detail")
    ws_detail.sheet_properties.tabColor = "10B981"

    detail_cols = [
        c for c in ["Dealer Name", "Distributor Name", "State", "Region",
                     "Shop Volume", "Site Volume", "Total Volume",
                     "Qualified Volume", "Qualified Slab", "Qualified Points"]
        if c in df.columns
    ]

    ws_detail.append(detail_cols)
    for _, row in df[detail_cols].iterrows():
        ws_detail.append(list(row.values))

    _style_header_row(ws_detail, len(detail_cols))
    _auto_width(ws_detail)
    ws_detail.freeze_panes = "A2"

    # Save
    wb.save(str(output_path))
    log_info(f"Report saved to {output_path}")
    return output_path


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------


def main() -> None:
    """Load data and generate the slab report."""
    try:
        df, month_cols = load_data()
        output = generate_report(df, month_cols)
        print(f"\nReport generated: {output}")
    except FileNotFoundError as e:
        log_error(str(e))
        print(f"\nError: {e}")
        print("Please add an .xlsx file to the data/ directory.")
        sys.exit(1)


if __name__ == "__main__":
    main()
