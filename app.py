"""
Streamlit Dashboard — Q4 Scheme Slab Analysis
==============================================
Single-file Streamlit app structured in numbered sections.
Place your Excel data file in the data/ directory.
Slabs are assigned based on Qualified Points from the source data.

Run: streamlit run app.py
"""

from __future__ import annotations

import glob
import os
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st

from src.lib.logger import error as log_error
from src.lib.logger import info as log_info

# ============================================================================
# SECTION 1 — CONSTANTS
# ============================================================================

DATA_DIR: str = "data"
SHEET_NAME: str = "Sheet1"
HEADER_ROW: int = 0

# Column name mapping: Excel column → internal standard name
COLUMN_MAP: dict[str, str] = {
    "Retailer Name": "Dealer Name",
    "State Name": "State",
    "District Name": "District",
    "Shop vol.": "Shop Volume",
    "Site vol.": "Site Volume",
    "Total vol. under scheme": "Qualified Volume",
    "Total site vol.": "Total Site Volume",
    "Points": "Qualified Points",
    "Current gift": "Gift",
    "Current gift slab": "Current Slab",
    "Distributor self-counter (Yes/No)": "Self Counter",
    "Jan + Feb+Mar": "Q4 Volume",
    "# Unique Site >200 MT": "Unique Site >200 MT",
}


def _find_excel_file() -> Path:
    """Locate the most recently modified Excel file in DATA_DIR.

    Returns:
        Path to the latest .xlsx file.

    Raises:
        SystemExit: Stops the Streamlit app if no file is found.
    """
    pattern = str(Path(DATA_DIR) / "*.xlsx")
    files = glob.glob(pattern)
    if not files:
        st.error(f"No .xlsx files found in `{DATA_DIR}/`. Please add your data file.")
        log_error(f"No Excel files in {DATA_DIR}/")
        st.stop()
    latest = max(files, key=os.path.getmtime)
    log_info(f"Using data file: {latest}")
    return Path(latest)


# ---------------------------------------------------------------------------
# SLAB_CONFIG — Single source of truth for all tier definitions
# ---------------------------------------------------------------------------
# Bounds are based on QUALIFIED POINTS (from the Excel data).

SLAB_CONFIG: list[dict] = [
    {
        "slab": "Unqualified",
        "slab_code": "-",
        "range": "0 – 749",
        "lower": 0,
        "upper": 750,
        "gift": "No Gift",
        "gift_full": "No Gift",
        "category": "Unqualified",
        "color": "#94a3b8",
    },
    {
        "slab": "Slab A",
        "slab_code": "A",
        "range": "750 – 2,999",
        "lower": 750,
        "upper": 3000,
        "gift": "Foot Massager",
        "gift_full": "Foot massager",
        "category": "A",
        "color": "#f59e0b",
    },
    {
        "slab": "Slab B",
        "slab_code": "B",
        "range": "3,000 – 4,199",
        "lower": 3000,
        "upper": 4200,
        "gift": "Sony Sound Bar",
        "gift_full": "Sony - Sound bar, woofer and speakers",
        "category": "B",
        "color": "#6366f1",
    },
    {
        "slab": "Slab C",
        "slab_code": "C",
        "range": "4,200 – 6,799",
        "lower": 4200,
        "upper": 6800,
        "gift": "Robot Vacuum",
        "gift_full": "Robot Vacuum cleaner",
        "category": "C",
        "color": "#10b981",
    },
    {
        "slab": "Slab D",
        "slab_code": "D",
        "range": "6,800 – 7,499",
        "lower": 6800,
        "upper": 7500,
        "gift": "Apple iPad",
        "gift_full": "Apple iPad",
        "category": "D",
        "color": "#3b82f6",
    },
    {
        "slab": "Slab E",
        "slab_code": "E",
        "range": "7,500+",
        "lower": 7500,
        "upper": float("inf"),
        "gift": "Washing Machine",
        "gift_full": "Samsung front-load washing machine",
        "category": "E",
        "color": "#ec4899",
    },
]

# Derived lookup dictionaries — all computed from SLAB_CONFIG
SLAB_GIFT_MAP: dict[str, str] = {s["slab"]: s["gift_full"] for s in SLAB_CONFIG}
SLAB_COLORS: dict[str, str] = {s["slab"]: s["color"] for s in SLAB_CONFIG}
SLAB_ORDER: list[str] = [s["slab"] for s in SLAB_CONFIG]

# Map Excel slab codes (A, B, C, ...) to full slab names
SLAB_CODE_MAP: dict[str, str] = {s["slab_code"]: s["slab"] for s in SLAB_CONFIG}

# Next-slab mapping
NEXT_SLAB_MAP: dict[str, Optional[str]] = {}
for _i, _cfg in enumerate(SLAB_CONFIG):
    NEXT_SLAB_MAP[_cfg["slab"]] = (
        SLAB_CONFIG[_i + 1]["slab"] if _i + 1 < len(SLAB_CONFIG) else None
    )

# Points threshold to reach next slab
NEXT_SLAB_THRESHOLD: dict[str, Optional[float]] = {}
for _i, _cfg in enumerate(SLAB_CONFIG):
    NEXT_SLAB_THRESHOLD[_cfg["slab"]] = (
        SLAB_CONFIG[_i + 1]["lower"] if _i + 1 < len(SLAB_CONFIG) else None
    )


# ============================================================================
# SECTION 2 — HELPER FUNCTIONS
# ============================================================================

def format_indian(number: float, prefix: str = "", decimal: int = 0) -> str:
    """Format a number with Indian comma grouping (e.g. 12,34,567).

    Args:
        number: The numeric value to format.
        prefix: Optional prefix string (e.g. "₹").
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

    # Indian grouping: last 3 digits, then groups of 2
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


def assign_slab(points: float) -> str:
    """Assign a slab label based on qualified points using SLAB_CONFIG.

    Args:
        points: Qualified points value.

    Returns:
        Slab label string (e.g. "Slab A").
    """
    if pd.isna(points):
        points = 0.0
    points = float(points)
    for cfg in SLAB_CONFIG:
        if cfg["lower"] <= points < cfg["upper"]:
            return cfg["slab"]
    return SLAB_CONFIG[0]["slab"]


def get_next_slab(current: str) -> Optional[str]:
    """Look up the next higher slab from NEXT_SLAB_MAP.

    Args:
        current: Current slab label.

    Returns:
        Next slab label, or None if already at maximum.
    """
    return NEXT_SLAB_MAP.get(current)


def points_to_next(points: float, current: str) -> Optional[float]:
    """Calculate the points gap to reach the next slab tier.

    Args:
        points: Current qualified points.
        current: Current slab label.

    Returns:
        Points needed to reach next tier, or None if at max slab.
    """
    threshold = NEXT_SLAB_THRESHOLD.get(current)
    if threshold is None:
        return None
    gap = threshold - float(points)
    return max(gap, 0.0)


# ============================================================================
# SECTION 3 — CUSTOM CSS
# ============================================================================

def inject_custom_css() -> None:
    """Inject custom CSS styles into the Streamlit app."""
    st.markdown(
        """
        <style>
        /* Header bar */
        .header-bar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 0.75rem 1rem;
            background: linear-gradient(135deg, #1e293b, #334155);
            border-radius: 0.5rem;
            margin-bottom: 1rem;
            color: #ffffff;
        }
        .header-bar h1 {
            margin: 0;
            font-size: 1.4rem;
            font-weight: 700;
            color: #ffffff;
        }
        .header-bar span {
            font-size: 0.85rem;
            color: #94a3b8;
        }

        /* Slab cards */
        .slab-card {
            background: linear-gradient(135deg, #f0f7ff, #e8f2ff);
            border-radius: 0.5rem;
            box-shadow: 0 1px 3px rgba(59,130,246,0.10);
            padding: 1rem 1.2rem;
            border-left: 4px solid #bfdbfe;
            margin-bottom: 0.5rem;
        }
        .slab-count {
            font-size: 1.7rem;
            font-weight: 700;
            color: #1e40af;
            line-height: 1.2;
        }
        .slab-label {
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #3b82f6;
            margin-top: 0.25rem;
        }
        .slab-gift {
            font-size: 0.72rem;
            color: #93c5fd;
            margin-top: 0.15rem;
        }

        /* KPI cards */
        .kpi-card {
            background: linear-gradient(135deg, #f8fbff, #eff6ff);
            border: 1px solid #dbeafe;
            border-radius: 0.5rem;
            box-shadow: 0 1px 3px rgba(59,130,246,0.08);
            padding: 1rem;
            text-align: center;
        }
        .kpi-label {
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #3b82f6;
            margin-bottom: 0.25rem;
        }
        .kpi-value {
            font-size: 1.35rem;
            font-weight: 700;
            color: #1e3a5f;
        }

        /* Total row styling for tables */
        .total-row {
            font-weight: 700;
            background-color: #eff6ff;
        }

        /* Light blue table styling */
        .stDataFrame [data-testid="stDataFrameResizable"] {
            border: 1px solid #bfdbfe;
            border-radius: 0.5rem;
        }
        .stDataFrame thead tr th {
            background-color: #dbeafe !important;
            color: #1e3a5f !important;
            font-weight: 600;
        }
        .stDataFrame tbody tr:nth-child(even) {
            background-color: #eff6ff !important;
        }
        .stDataFrame tbody tr:nth-child(odd) {
            background-color: #f8fbff !important;
        }
        .stDataFrame tbody tr:hover {
            background-color: #dbeafe !important;
        }

        /* Hide Streamlit chrome */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        /* Streamlit tab styling */
        .stTabs [data-baseweb="tab-list"] {
            gap: 0.5rem;
        }
        .stTabs [data-baseweb="tab"] {
            padding: 0.5rem 1rem;
            font-size: 0.85rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================================
# SECTION 4 — DATA LOADING
# ============================================================================

@st.cache_data
def load_data() -> pd.DataFrame:
    """Load, clean, and return the Excel data.

    Reads the latest .xlsx from DATA_DIR, renames columns to standard names,
    excludes self-counter dealers, and derives the Qualified Slab from the
    pre-calculated Points column in the Excel.

    Returns:
        Cleaned DataFrame.
    """
    file_path = _find_excel_file()
    log_info(f"Loading data from {file_path}")

    try:
        df = pd.read_excel(
            file_path,
            sheet_name=SHEET_NAME,
            header=HEADER_ROW,
            engine="openpyxl",
        )
    except FileNotFoundError:
        st.error(f"File not found: {file_path}")
        log_error(f"File not found: {file_path}")
        st.stop()
    except ValueError as e:
        st.error(f"Error reading Excel sheet '{SHEET_NAME}': {e}")
        log_error(f"Sheet error: {e}")
        st.stop()

    log_info(f"Loaded {len(df)} rows, {len(df.columns)} columns")

    # --- Rename columns to standard names ---
    rename_map = {k: v for k, v in COLUMN_MAP.items() if k in df.columns}
    df = df.rename(columns=rename_map)
    log_info(f"Renamed columns: {list(rename_map.values())}")

    # --- Exclude self-counter dealers ---
    if "Self Counter" in df.columns:
        before = len(df)
        df["Self Counter"] = df["Self Counter"].fillna("").astype(str).str.strip().str.title()
        df = df[df["Self Counter"] != "Yes"].copy()
        excluded = before - len(df)
        log_info(f"Excluded {excluded} self-counter dealers ({len(df)} remaining)")

    # --- Coerce numeric columns ---
    numeric_cols = [
        "Qualified Volume", "Total Site Volume", "Qualified Points",
        "Shop Volume", "Site Volume", "Q4 Volume", "Unique Site >200 MT",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # --- Clean text columns ---
    text_cols = ["Dealer Name", "Distributor Name", "State", "District", "Zone"]
    for col in text_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.title()
                .replace({"Nan": "", "None": "", "0": "", "0.0": ""})
            )

    # --- Total Volume (Shop + Site as raw total) ---
    if "Shop Volume" in df.columns and "Site Volume" in df.columns:
        df["Total Volume"] = df["Shop Volume"] + df["Site Volume"]
    elif "Q4 Volume" in df.columns:
        df["Total Volume"] = df["Q4 Volume"]
    else:
        df["Total Volume"] = 0.0

    # --- Derive Qualified Slab from Points ---
    if "Qualified Points" in df.columns:
        df["Qualified Slab"] = df["Qualified Points"].apply(assign_slab)
    elif "Current Slab" in df.columns:
        # Fallback: map from Excel slab code
        df["Current Slab"] = df["Current Slab"].fillna("-").astype(str).str.strip()
        df["Qualified Slab"] = df["Current Slab"].map(SLAB_CODE_MAP).fillna("Unqualified")
    else:
        df["Qualified Slab"] = "Unqualified"

    # --- Next slab & points gap ---
    df["Next Upgrade Slab"] = df["Qualified Slab"].apply(get_next_slab)
    if "Qualified Points" in df.columns:
        df["Points to Next Slab"] = df.apply(
            lambda row: points_to_next(row["Qualified Points"], row["Qualified Slab"]),
            axis=1,
        )

    log_info(f"Data loading complete. Shape: {df.shape}")
    return df


# ============================================================================
# SECTION 5 — CASCADING FILTERS
# ============================================================================

def _opts(series: pd.Series) -> list[str]:
    """Build filter options: ['All'] + sorted unique non-blank values.

    Args:
        series: Pandas Series to extract unique values from.

    Returns:
        List starting with 'All', followed by sorted unique values.
    """
    unique_vals = series[series.astype(str).str.strip() != ""].unique()
    return ["All"] + sorted(str(v) for v in unique_vals if str(v).strip())


def render_cascading_filters(
    df: pd.DataFrame,
    key: str,
    filter_fields: Optional[list[str]] = None,
) -> pd.DataFrame:
    """Render cascading dropdown filters and return the filtered DataFrame.

    Args:
        df: Input DataFrame to filter.
        key: Unique key prefix for widget state (use different keys per tab).
        filter_fields: Optional explicit list of column names to filter on.

    Returns:
        Filtered DataFrame based on user selections.
    """
    if filter_fields is None:
        filter_fields = [
            col for col in ["Qualified Slab", "Zone", "State", "District", "Distributor Name"]
            if col in df.columns
        ]

    filter_fields = [f for f in filter_fields if f in df.columns]

    if not filter_fields:
        return df

    cols = st.columns(len(filter_fields))
    filtered = df.copy()

    for i, field in enumerate(filter_fields):
        with cols[i]:
            options = _opts(filtered[field])
            selected = st.selectbox(
                field,
                options,
                key=f"{key}_{field}",
            )
            if selected != "All":
                filtered = filtered[filtered[field] == selected]

    return filtered


# ============================================================================
# SECTION 6 — TAB: SUMMARY
# ============================================================================

def render_summary_top(df: pd.DataFrame) -> None:
    """Render the top summary section with KPIs and slab cards (no filters).

    Args:
        df: Full DataFrame.
    """
    # --- KPI Row ---
    total_dealers = len(df)
    total_volume = df["Total Volume"].sum() if "Total Volume" in df.columns else 0
    total_shop_vol = df["Shop Volume"].sum() if "Shop Volume" in df.columns else 0
    total_site_vol = df["Site Volume"].sum() if "Site Volume" in df.columns else 0
    total_points = df["Qualified Points"].sum() if "Qualified Points" in df.columns else 0

    kpi_cols = st.columns(5)
    kpis = [
        ("Total Dealers", format_indian(total_dealers)),
        ("Total Volume (MT)", format_indian(total_volume, decimal=1)),
        ("Shop Qual. Volume (MT)", format_indian(total_shop_vol, decimal=1)),
        ("Site Qual. Volume (MT)", format_indian(total_site_vol, decimal=1)),
        ("Total Qualified Points", format_indian(total_points, decimal=1)),
    ]
    for col, (label, value) in zip(kpi_cols, kpis):
        with col:
            st.markdown(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">{label}</div>
                    <div class="kpi-value">{value}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)

    # --- Slab Distribution Cards ---
    if "Qualified Slab" in df.columns:
        slab_counts = df["Qualified Slab"].value_counts()
        card_cols = st.columns(len(SLAB_CONFIG))
        for col, cfg in zip(card_cols, SLAB_CONFIG):
            count = int(slab_counts.get(cfg["slab"], 0))
            color = cfg["color"]
            with col:
                st.markdown(
                    f"""
                    <div class="slab-card" style="border-left-color: {color};">
                        <div class="slab-count" style="color: {color};">{count}</div>
                        <div class="slab-label">{cfg['slab']} — {cfg['range']} pts</div>
                        <div class="slab-gift">{cfg['gift']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    st.markdown("<br>", unsafe_allow_html=True)


def render_summary(df: pd.DataFrame) -> None:
    """Render the Summary tab with breakdown table and chart (no filters).

    Args:
        df: Full DataFrame.
    """
    # --- Slab Breakdown Table ---
    st.subheader("Slab Breakdown")
    summary_rows = []
    grand_count = 0
    grand_vol = 0.0
    grand_shop_vol = 0.0
    grand_site_vol = 0.0
    grand_pts = 0.0

    for cfg in SLAB_CONFIG:
        slab_df = df[df["Qualified Slab"] == cfg["slab"]]
        count = len(slab_df)
        vol = slab_df["Total Volume"].sum() if "Total Volume" in slab_df.columns else 0
        shop_vol = slab_df["Shop Volume"].sum() if "Shop Volume" in slab_df.columns else 0
        site_vol = slab_df["Site Volume"].sum() if "Site Volume" in slab_df.columns else 0
        pts = slab_df["Qualified Points"].sum() if "Qualified Points" in slab_df.columns else 0
        grand_count += count
        grand_vol += vol
        grand_shop_vol += shop_vol
        grand_site_vol += site_vol
        grand_pts += pts
        summary_rows.append({
            "Slab": cfg["slab"],
            "Points Range": cfg["range"],
            "Dealer Count": count,
            "Total Volume": format_indian(vol, decimal=1),
            "Shop Qual. Volume": format_indian(shop_vol, decimal=1),
            "Site Qual. Volume": format_indian(site_vol, decimal=1),
            "Total Points": format_indian(pts, decimal=1),
            "Gift": cfg["gift_full"],
        })

    summary_rows.append({
        "Slab": "TOTAL",
        "Points Range": "",
        "Dealer Count": grand_count,
        "Total Volume": format_indian(grand_vol, decimal=1),
        "Shop Qual. Volume": format_indian(grand_shop_vol, decimal=1),
        "Site Qual. Volume": format_indian(grand_site_vol, decimal=1),
        "Total Points": format_indian(grand_pts, decimal=1),
        "Gift": "",
    })

    summary_df = pd.DataFrame(summary_rows)
    st.dataframe(summary_df, width="stretch", hide_index=True)


# ============================================================================
# SECTION 7 — TAB: DEALER DETAILS
# ============================================================================

def render_dealer_details(df: pd.DataFrame) -> None:
    """Render the Dealer Details tab with per-dealer table and filters.

    Filters: Distributor Name, State, Dealer Name, Qualified Slab.
    Columns: Dealer Name, Distributor Name, State, Region/Zone,
             Shop Volume, Site Volume, Total Volume, Qualified Volume,
             Qualified Slab, Qualified Points.

    Args:
        df: Full DataFrame.
    """
    detail_filters = ["Distributor Name", "State", "Dealer Name", "Qualified Slab"]
    filtered = render_cascading_filters(df, key="dealer_detail", filter_fields=detail_filters)

    if filtered.empty:
        st.info("No data matches the selected filters.")
        return

    # --- Data Table ---
    display_cols = [
        c for c in [
            "Dealer Name", "Distributor Name", "State", "Zone",
            "Shop Volume", "Site Volume", "Total Volume",
            "Qualified Slab", "Qualified Points",
        ]
        if c in filtered.columns
    ]

    st.subheader(f"Dealer Details ({len(filtered)} records)")

    # Round numeric columns to 1 decimal
    display_df = filtered[display_cols].copy()
    num_cols = ["Shop Volume", "Site Volume", "Total Volume", "Qualified Points"]
    for col in num_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].round(1)

    def _highlight_total_row(row: pd.Series) -> list[str]:
        """Apply bold grey background to total/summary rows."""
        for field in ["Dealer Name", "Distributor Name"]:
            val = row.get(field)
            if isinstance(val, str) and "total" in val.lower():
                return ["font-weight: 700; background-color: #f1f5f9"] * len(row)
        return [""] * len(row)

    styled = display_df.style.apply(_highlight_total_row, axis=1)
    st.dataframe(styled, width="stretch", hide_index=True, height=500)


# ============================================================================
# SECTION 8 — MAIN
# ============================================================================

def main() -> None:
    """Entry point for the Streamlit dashboard application."""
    st.set_page_config(
        page_title="Q4 Scheme Dashboard",
        layout="wide",
        page_icon="📊",
    )

    inject_custom_css()

    # --- Header Bar ---
    st.markdown(
        """
        <div class="header-bar">
            <h1>📊 Q4 Scheme Dashboard</h1>
            <span>Slab Analysis & Dealer Tracker</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --- Load Data ---
    df = load_data()

    # --- File info ---
    file_path = _find_excel_file()
    st.caption(f"Data source: `{file_path.name}` — {len(df)} dealers loaded (excl. self-counter)")

    # --- Summary on top (always visible, no filters) ---
    render_summary_top(df)

    # --- Tabs ---
    tab_summary, tab_details = st.tabs([
        "📊 Summary",
        "🔍 Dealer Details",
    ])

    with tab_summary:
        render_summary(df)

    with tab_details:
        render_dealer_details(df)


if __name__ == "__main__":
    main()
