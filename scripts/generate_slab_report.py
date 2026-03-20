"""
Standalone Slab Report Generator
================================
Generates a formatted Excel report with slab analysis.
Outputs to .workspace/ directory.

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
    {"slab": "Slab 1", "range": "0 – 499", "lower": 0, "upper": 500,
     "gift": "No Gift", "gift_full": "Below minimum qualification",
     "category": "Unqualified", "value": 0, "value_tds": 0, "color": "#94a3b8"},
    {"slab": "Slab 2", "range": "500 – 999", "lower": 500, "upper": 1000,
     "gift": "Bronze Gift", "gift_full": "Bronze tier reward",
     "category": "Bronze", "value": 5000, "value_tds": 4500, "color": "#f59e0b"},
    {"slab": "Slab 3", "range": "1,000 – 2,499", "lower": 1000, "upper": 2500,
     "gift": "Silver Gift", "gift_full": "Silver tier reward",
     "category": "Silver", "value": 15000, "value_tds": 13500, "color": "#6366f1"},
    {"slab": "Slab 4", "range": "2,500 – 4,999", "lower": 2500, "upper": 5000,
     "gift": "Gold Gift", "gift_full": "Gold tier reward",
     "category": "Gold", "value": 35000, "value_tds": 31500, "color": "#10b981"},
    {"slab": "Slab 5", "range": "5,000+", "lower": 5000, "upper": float("inf"),
     "gift": "Platinum Gift", "gift_full": "Platinum tier reward",
     "category": "Platinum", "value": 75000, "value_tds": 67500, "color": "#3b82f6"},
]

SLAB_ORDER: list[str] = [s["slab"] for s in SLAB_CONFIG]

# ---------------------------------------------------------------------------
# HELPERS (mirrored from app.py)
# ---------------------------------------------------------------------------


def assign_slab(vol: float) -> str:
    """Assign a slab label based on volume.

    Args:
        vol: The volume value to classify.

    Returns:
        Slab label string.
    """
    if pd.isna(vol):
        vol = 0.0
    vol = float(vol)
    for cfg in SLAB_CONFIG:
        if cfg["lower"] <= vol < cfg["upper"]:
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

    # Derived columns
    if month_cols:
        if "Total Volume" not in df.columns:
            df["Total Volume"] = df[month_cols].sum(axis=1)
        df["Qualified Slab"] = df["Total Volume"].apply(assign_slab)

    log_info(f"Loaded {len(df)} rows")
    return df, month_cols


# ---------------------------------------------------------------------------
# EXCEL REPORT GENERATION
# ---------------------------------------------------------------------------

# Styles
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

    headers = ["Slab", "Range", "Count", "Total Volume", "Gift", "Value (post-TDS)"]
    ws_summary.append(headers)

    for cfg in SLAB_CONFIG:
        slab_df = df[df["Qualified Slab"] == cfg["slab"]] if "Qualified Slab" in df.columns else pd.DataFrame()
        count = len(slab_df)
        vol = slab_df["Total Volume"].sum() if "Total Volume" in slab_df.columns else 0
        ws_summary.append([
            cfg["slab"],
            cfg["range"],
            count,
            round(vol, 2),
            cfg["gift_full"],
            cfg["value_tds"] * count,
        ])

    _style_header_row(ws_summary, len(headers))
    _auto_width(ws_summary)
    ws_summary.freeze_panes = "A2"

    # --- Sheet 2: Distributor Detail ---
    ws_detail = wb.create_sheet("Distributor Detail")
    ws_detail.sheet_properties.tabColor = "10B981"

    detail_cols = [
        c for c in ["Distributor Name", "State", "District", "Zone",
                     "Total Volume", "Qualified Slab"] + month_cols
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
