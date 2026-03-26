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
DATA_FILE: str = "Q4 as on 25th Mar.xlsx"
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
    """Locate the configured Excel data file in DATA_DIR.

    Uses the explicit DATA_FILE constant if the file exists, otherwise
    falls back to the most recently modified .xlsx file.

    Returns:
        Path to the data .xlsx file.

    Raises:
        SystemExit: Stops the Streamlit app if no file is found.
    """
    explicit = Path(DATA_DIR) / DATA_FILE
    if explicit.exists():
        log_info(f"Using data file: {explicit}")
        return explicit

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
        "volume_mt": 0,
        "color": "#94a3b8",
        "color_light": "#f1f5f9",
        "gift_inr": 0,
        "threshold_points": 0,
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
        "volume_mt": 30,
        "color": "#f59e0b",
        "color_light": "#fef3c7",
        "gift_inr": 3750,
        "threshold_points": 750,
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
        "volume_mt": 120,
        "color": "#6366f1",
        "color_light": "#e0e7ff",
        "gift_inr": 15000,
        "threshold_points": 3000,
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
        "volume_mt": 168,
        "color": "#10b981",
        "color_light": "#d1fae5",
        "gift_inr": 21000,
        "threshold_points": 4200,
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
        "volume_mt": 272,
        "color": "#3b82f6",
        "color_light": "#dbeafe",
        "gift_inr": 34000,
        "threshold_points": 6800,
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
        "volume_mt": 300,
        "color": "#ec4899",
        "color_light": "#fce7f3",
        "gift_inr": 37500,
        "threshold_points": 7500,
    },
]

# Derived lookup dictionaries — all computed from SLAB_CONFIG
SLAB_GIFT_MAP: dict[str, str] = {s["slab"]: s["gift_full"] for s in SLAB_CONFIG}
SLAB_COLORS: dict[str, str] = {s["slab"]: s["color"] for s in SLAB_CONFIG}
SLAB_COLORS_LIGHT: dict[str, str] = {s["slab"]: s["color_light"] for s in SLAB_CONFIG}
SLAB_ORDER: list[str] = [s["slab"] for s in SLAB_CONFIG]
SLAB_GIFT_INR: dict[str, int] = {s["slab"]: s["gift_inr"] for s in SLAB_CONFIG}
SLAB_THRESHOLD_PTS: dict[str, int] = {s["slab"]: s["threshold_points"] for s in SLAB_CONFIG}

TOTAL_RETAIL_SALES: float = 42126.0

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

# Volume (MT) threshold to reach next slab
NEXT_SLAB_VOLUME: dict[str, Optional[float]] = {}
for _i, _cfg in enumerate(SLAB_CONFIG):
    NEXT_SLAB_VOLUME[_cfg["slab"]] = (
        SLAB_CONFIG[_i + 1]["volume_mt"] if _i + 1 < len(SLAB_CONFIG) else None
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
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        /* Global page styling */
        .stApp {
            background-color: #f8fafc;
        }
        .block-container {
            padding-top: 1rem !important;
        }

        /* ── Header bar ── */
        .header-bar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 1rem 1.5rem;
            background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 50%, #1e40af 100%);
            border-radius: 0.75rem;
            margin-bottom: 0.75rem;
            color: #ffffff;
            box-shadow: 0 4px 12px rgba(15,23,42,0.25);
        }
        .header-bar h1 {
            margin: 0;
            font-size: 1.5rem;
            font-weight: 800;
            color: #ffffff;
            letter-spacing: -0.01em;
        }
        .header-bar .header-right {
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }
        .header-bar .header-badge {
            background: rgba(255,255,255,0.15);
            border: 1px solid rgba(255,255,255,0.25);
            border-radius: 2rem;
            padding: 0.3rem 0.9rem;
            font-size: 0.8rem;
            font-weight: 600;
            color: #e2e8f0;
            letter-spacing: 0.03em;
        }
        .header-bar .header-subtitle {
            font-size: 0.85rem;
            font-weight: 500;
            color: #94a3b8;
        }
        .header-badge {
            font-size: 0.85rem;
            font-weight: 600;
            color: #ffffff;
            background: rgba(255, 255, 255, 0.15);
            border: 1px solid rgba(255, 255, 255, 0.3);
            border-radius: 1rem;
            padding: 0.3rem 1rem;
        }

        /* ── Section divider ── */
        .section-divider {
            border: none;
            border-top: 1px solid #e2e8f0;
            margin: 0.5rem 0 1rem 0;
        }
        .section-title {
            font-size: 0.8rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: #64748b;
            margin-bottom: 0.5rem;
        }

        /* ── Filter panel ── */
        .filter-panel {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 0.75rem;
            padding: 0.9rem 1.2rem 0.5rem 1.2rem;
            margin-bottom: 1rem;
            box-shadow: 0 1px 4px rgba(0,0,0,0.04);
        }
        .filter-panel-title {
            font-size: 0.75rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.1em;
            color: #64748b;
            margin-bottom: 0.4rem;
            display: flex;
            align-items: center;
            gap: 0.4rem;
        }

        /* ── KPI cards ── */
        .kpi-card {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-top: 3px solid #2563eb;
            border-radius: 0.6rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            padding: 1rem 0.8rem;
            text-align: center;
            transition: box-shadow 0.2s, transform 0.2s;
        }
        .kpi-card:hover {
            box-shadow: 0 4px 16px rgba(37,99,235,0.12);
            transform: translateY(-1px);
        }
        .kpi-label {
            font-size: 0.72rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            color: #64748b;
            margin-bottom: 0.35rem;
        }
        .kpi-value {
            font-size: 1.45rem;
            font-weight: 800;
            color: #0f172a;
            letter-spacing: -0.02em;
        }

        /* ── Slab cards ── */
        .slab-card {
            background: #ffffff;
            border-radius: 0.6rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.07);
            padding: 0;
            margin-bottom: 0.5rem;
            overflow: hidden;
            border: 1px solid #e2e8f0;
            transition: box-shadow 0.2s, transform 0.2s;
        }
        .slab-card:hover {
            box-shadow: 0 4px 16px rgba(0,0,0,0.12);
            transform: translateY(-1px);
        }
        .slab-color-band {
            height: 5px;
            width: 100%;
        }
        .slab-card-body {
            padding: 0.9rem 1rem;
        }
        .slab-count {
            font-size: 2rem;
            font-weight: 800;
            color: #0f172a;
            line-height: 1.1;
        }
        .slab-pct {
            font-size: 0.75rem;
            font-weight: 600;
            color: #64748b;
            margin-left: 0.3rem;
        }
        .slab-label {
            font-size: 0.8rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            color: #1e293b;
            margin-top: 0.25rem;
        }
        .slab-gift {
            font-size: 0.73rem;
            font-weight: 500;
            color: #64748b;
            margin-top: 0.15rem;
            font-style: italic;
        }

        /* ── Total row styling for tables ── */
        .total-row {
            font-weight: 800;
            background-color: #e2e8f0;
        }

        /* ── Table styling ── */
        .stDataFrame [data-testid="stDataFrameResizable"] {
            border: 1px solid #cbd5e1;
            border-radius: 0.6rem;
            overflow: hidden;
        }
        .stDataFrame thead tr th {
            background-color: #1e293b !important;
            color: #ffffff !important;
            font-weight: 700 !important;
            font-size: 0.82rem !important;
            letter-spacing: 0.02em !important;
        }
        /* Bold headers in Streamlit dataframe (glide-data-grid) */
        [data-testid="stDataFrame"] [role="columnheader"],
        [data-testid="stDataFrame"] [data-testid="glide-cell"] {
            font-weight: 700 !important;
        }
        [data-testid="stDataFrame"] .gdg-header {
            font-weight: 800 !important;
        }
        .stDataFrame tbody tr:nth-child(even) {
            background-color: #f8fafc !important;
        }
        .stDataFrame tbody tr:nth-child(odd) {
            background-color: #ffffff !important;
        }
        .stDataFrame tbody tr:hover {
            background-color: #e2e8f0 !important;
        }

        /* ── Hide Streamlit chrome ── */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        /* ── Streamlit tab styling ── */
        .stTabs [data-baseweb="tab-list"] {
            gap: 0;
            background: #ffffff;
            border-radius: 0.6rem 0.6rem 0 0;
            border-bottom: 2px solid #e2e8f0;
            padding: 0 0.5rem;
        }
        .stTabs [data-baseweb="tab"] {
            padding: 0.7rem 1.5rem;
            font-size: 0.88rem;
            font-weight: 600;
            color: #64748b;
            border-bottom: 3px solid transparent;
            transition: color 0.2s, border-color 0.2s;
        }
        .stTabs [data-baseweb="tab"]:hover {
            color: #334155;
        }
        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            font-weight: 700;
            color: #1e40af;
            border-bottom: 3px solid #2563eb;
        }

        /* ── Filter dropdowns ── */
        .stSelectbox label {
            font-weight: 600 !important;
            color: #334155 !important;
            font-size: 0.8rem !important;
        }
        .stSelectbox [data-baseweb="select"] {
            border-color: #cbd5e1 !important;
            border-radius: 0.4rem !important;
        }

        /* ── Subheaders ── */
        h2, h3 {
            color: #0f172a !important;
            font-weight: 700 !important;
        }

        /* ── Caption / data source ── */
        .stCaption, [data-testid="stCaptionContainer"] {
            color: #94a3b8 !important;
            font-size: 0.75rem !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================================
# SECTION 4 — DATA LOADING
# ============================================================================

@st.cache_data(ttl=3600)
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

    # --- Total Volume (Jan + Feb + Mar) ---
    if "Q4 Volume" in df.columns:
        df["Total Volume"] = df["Q4 Volume"]
    elif "Shop Volume" in df.columns and "Total Site Volume" in df.columns:
        df["Total Volume"] = df["Shop Volume"] + df["Total Site Volume"]
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
    """Render the top summary section with KPIs and slab cards.

    Args:
        df: Filtered DataFrame.
    """
    # --- KPI Row ---
    st.markdown('<div class="section-title">Key Metrics</div>', unsafe_allow_html=True)
    total_dealers = len(df)
    total_volume = df["Total Volume"].sum() if "Total Volume" in df.columns else 0
    total_points = df["Qualified Points"].sum() if "Qualified Points" in df.columns else 0

    total_shop_vol = df["Shop Volume"].sum() if "Shop Volume" in df.columns else 0
    total_site_vol = df["Total Site Volume"].sum() if "Total Site Volume" in df.columns else 0
    total_qual_vol = df["Qualified Volume"].sum() if "Qualified Volume" in df.columns else 0

    kpi_cols = st.columns(6)
    kpis = [
        ("Total Dealers", format_indian(total_dealers)),
        ("Total Volume (MT)", format_indian(total_volume, decimal=0)),
        ("Qual. Shop Vol. (MT)", format_indian(total_shop_vol, decimal=0)),
        ("Qual. Site Vol. (MT)", format_indian(total_site_vol, decimal=0)),
        ("Qualified Volume (MT)", format_indian(total_qual_vol, decimal=0)),
        ("Total Qualified Points", format_indian(total_points, decimal=0)),
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

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # --- Slab Distribution Cards ---
    if "Qualified Slab" in df.columns:
        total_count = len(df)
        slab_counts = df["Qualified Slab"].value_counts()
        st.markdown('<div class="section-title">Slab Distribution</div>', unsafe_allow_html=True)
        card_cols = st.columns(len(SLAB_CONFIG))
        for col, cfg in zip(card_cols, SLAB_CONFIG):
            count = int(slab_counts.get(cfg["slab"], 0))
            pct = f"{count / total_count * 100:.0f}%" if total_count > 0 else "0%"
            color = cfg["color"]
            with col:
                st.markdown(
                    f"""
                    <div class="slab-card">
                        <div class="slab-color-band" style="background: {color};"></div>
                        <div class="slab-card-body">
                            <div class="slab-count" style="color: {color};">
                                {count}<span class="slab-pct">({pct})</span>
                            </div>
                            <div class="slab-label">{cfg['slab']} — {cfg['range']} pts</div>
                            <div class="slab-gift">{cfg['gift']}</div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)


def render_summary(df: pd.DataFrame) -> None:
    """Render the Summary tab with slab breakdown table.

    Args:
        df: Filtered DataFrame.
    """
    # --- Slab Breakdown Table ---
    st.subheader("Slab Breakdown")
    summary_rows = []
    grand_count = 0
    grand_vol = 0.0
    grand_shop_vol = 0.0
    grand_site_vol = 0.0
    grand_qual_vol = 0.0
    grand_pts = 0.0

    for cfg in SLAB_CONFIG:
        slab_df = df[df["Qualified Slab"] == cfg["slab"]]
        count = len(slab_df)
        vol = slab_df["Total Volume"].sum() if "Total Volume" in slab_df.columns else 0
        shop_vol = slab_df["Shop Volume"].sum() if "Shop Volume" in slab_df.columns else 0
        site_vol = slab_df["Total Site Volume"].sum() if "Total Site Volume" in slab_df.columns else 0
        qual_vol = slab_df["Qualified Volume"].sum() if "Qualified Volume" in slab_df.columns else 0
        pts = slab_df["Qualified Points"].sum() if "Qualified Points" in slab_df.columns else 0
        grand_count += count
        grand_vol += vol
        grand_shop_vol += shop_vol
        grand_site_vol += site_vol
        grand_qual_vol += qual_vol
        grand_pts += pts
        summary_rows.append({
            "Slab": cfg["slab"],
            "Points Range": cfg["range"],
            "Dealer Count": count,
            "Total Volume": format_indian(vol, decimal=0),
            "Qual. Shop Vol.": format_indian(shop_vol, decimal=0),
            "Qual. Site Vol.": format_indian(site_vol, decimal=0),
            "Qualified Volume": format_indian(qual_vol, decimal=0),
            "Total Points": format_indian(pts, decimal=0),
            "Gift": cfg["gift_full"],
        })

    summary_rows.append({
        "Slab": "TOTAL",
        "Points Range": "",
        "Dealer Count": grand_count,
        "Total Volume": format_indian(grand_vol, decimal=0),
        "Qual. Shop Vol.": format_indian(grand_shop_vol, decimal=0),
        "Qual. Site Vol.": format_indian(grand_site_vol, decimal=0),
        "Qualified Volume": format_indian(grand_qual_vol, decimal=0),
        "Total Points": format_indian(grand_pts, decimal=0),
        "Gift": "",
    })

    summary_df = pd.DataFrame(summary_rows)

    def _style_summary_row(row: pd.Series) -> list[str]:
        """Apply slab color and bold formatting to summary rows."""
        slab = row.get("Slab", "")
        if slab == "TOTAL":
            return ["font-weight: 800; background-color: #cbd5e1; color: #0f172a"] * len(row)
        bg = SLAB_COLORS_LIGHT.get(slab, "")
        if bg:
            return [f"background-color: {bg}"] * len(row)
        return [""] * len(row)

    def _bold_summary_columns(col: pd.Series) -> list[str]:
        """Bold key summary columns."""
        if col.name in ("Slab", "Dealer Count", "Total Points"):
            return ["font-weight: 700"] * len(col)
        return [""] * len(col)

    styled = (
        summary_df.style
        .apply(_style_summary_row, axis=1)
        .apply(_bold_summary_columns, axis=0)
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)


# ============================================================================
# SECTION 7 — TAB: DEALER DETAILS
# ============================================================================

def _calc_vol_to_achieve(slab: str, total_volume: float) -> float:
    """Calculate volume remaining to reach the next slab tier.

    Args:
        slab: Current qualified slab label.
        total_volume: Dealer's current total volume.

    Returns:
        Volume gap to next slab, or 0 if at max slab.
    """
    next_vol = NEXT_SLAB_VOLUME.get(slab)
    if next_vol is None:
        return 0
    vol = float(total_volume) if not pd.isna(total_volume) else 0.0
    gap = next_vol - vol
    return max(gap, 0)


def render_dealer_details(df: pd.DataFrame) -> None:
    """Render the Dealer Details tab with per-dealer table and filters.

    Filters: Distributor Name, State, Dealer Name, Qualified Slab.
    Columns: Dealer Name, Distributor Name, State, Zone,
             Qual. Shop Vol., Qual. Site Vol., Total Volume,
             Qualified Slab, Next Slab, Vol. to Achieve, Qualified Points.

    Args:
        df: Full DataFrame.
    """
    # --- All 5 filters in one row: 4 cascading + Vol. to Achieve ---
    detail_filters = ["Distributor Name", "State", "Dealer Name", "Qualified Slab"]
    _all_cols = st.columns(len(detail_filters) + 1)
    filtered = df.copy()

    for i, field in enumerate(detail_filters):
        with _all_cols[i]:
            options = _opts(filtered[field])
            selected = st.selectbox(
                field,
                options,
                key=f"dealer_detail_{field}",
            )
            if selected != "All":
                filtered = filtered[filtered[field] == selected]

    _vol_options = ["All", "0 - 20 MT", "21 - 40 MT", "41 - 60 MT", "61 - 100 MT", "More than 100 MT"]
    with _all_cols[-1]:
        _vol_sel = st.selectbox("Vol. to Achieve", _vol_options, key="dealer_detail_vol_achieve")

    if _vol_sel != "All":
        _temp_vol = filtered.apply(
            lambda row: _calc_vol_to_achieve(
                row.get("Qualified Slab", ""),
                row.get("Total Volume", 0),
            ),
            axis=1,
        )
        _range_map: dict[str, tuple[float, float]] = {
            "0 - 20 MT": (0, 20),
            "21 - 40 MT": (21, 40),
            "41 - 60 MT": (41, 60),
            "61 - 100 MT": (61, 100),
            "More than 100 MT": (101, float("inf")),
        }
        _lo, _hi = _range_map[_vol_sel]
        filtered = filtered[(_temp_vol >= _lo) & (_temp_vol <= _hi)]

    if filtered.empty:
        st.info("No data matches the selected filters.")
        return

    # --- Build display DataFrame ---
    source_cols = [
        c for c in [
            "Dealer Name", "Distributor Name", "State", "Zone",
            "Shop Volume", "Total Site Volume", "Total Volume",
            "Qualified Slab", "Next Upgrade Slab",
            "Qualified Points",
        ]
        if c in filtered.columns
    ]

    st.subheader(f"Dealer Details ({len(filtered)} records)")

    display_df = filtered[source_cols].copy()

    # Rename columns for display
    rename_map: dict[str, str] = {
        "Shop Volume": "Qual. Shop Vol.",
        "Total Site Volume": "Qual. Site Vol.",
        "Next Upgrade Slab": "Next Slab",
    }
    display_df = display_df.rename(columns=rename_map)

    # Fill None values for Next Slab (Slab E dealers)
    if "Next Slab" in display_df.columns:
        display_df["Next Slab"] = display_df["Next Slab"].fillna("-")

    # Compute Vol. to Achieve = next slab volume threshold - dealer's Total Volume
    display_df["Vol. to Achieve"] = display_df.apply(
        lambda row: _calc_vol_to_achieve(
            row.get("Qualified Slab", ""),
            row.get("Total Volume", 0),
        ),
        axis=1,
    )

    # Reorder columns so Vol. to Achieve appears before Qualified Points
    ordered_cols = [
        c for c in [
            "Dealer Name", "Distributor Name", "State", "Zone",
            "Qual. Shop Vol.", "Qual. Site Vol.", "Total Volume",
            "Qualified Slab", "Next Slab", "Vol. to Achieve", "Qualified Points",
        ]
        if c in display_df.columns
    ]
    display_df = display_df[ordered_cols]

    # Round all numeric columns to whole numbers
    num_cols = [
        "Qual. Shop Vol.", "Qual. Site Vol.", "Total Volume",
        "Vol. to Achieve", "Qualified Points",
    ]
    for col in num_cols:
        if col in display_df.columns:
            display_df[col] = pd.to_numeric(
                display_df[col], errors="coerce"
            ).fillna(0).round(0).astype(int)

    def _highlight_by_slab(row: pd.Series) -> list[str]:
        """Apply light slab-based background color to each row."""
        for field in ["Dealer Name", "Distributor Name"]:
            val = row.get(field)
            if isinstance(val, str) and "total" in val.lower():
                return ["font-weight: 700; background-color: #f1f5f9"] * len(row)

        slab = row.get("Qualified Slab", "")
        bg_color = SLAB_COLORS_LIGHT.get(slab, "")
        if bg_color:
            return [f"background-color: {bg_color}"] * len(row)
        return [""] * len(row)

    def _bold_key_columns(col: pd.Series) -> list[str]:
        """Bold key columns: Dealer Name, Qualified Slab, Qualified Points."""
        if col.name in ("Dealer Name", "Qualified Slab", "Qualified Points",
                         "Next Slab", "Vol. to Achieve"):
            return ["font-weight: 700"] * len(col)
        return [""] * len(col)

    styled = (
        display_df.style
        .apply(_highlight_by_slab, axis=1)
        .apply(_bold_key_columns, axis=0)
    )
    st.dataframe(styled, use_container_width=True, hide_index=True, height=500)


# ============================================================================
# SECTION 8 — TAB: NEAR-UPGRADE ANALYSIS
# ============================================================================

def render_near_upgrade(df: pd.DataFrame) -> None:
    """Render the Near-Upgrade Analysis tab showing dealers close to next slab.

    Displays dealers within a configurable points threshold of upgrading,
    sorted by points remaining (ascending) so the sales team can prioritize.

    Args:
        df: Filtered DataFrame.
    """
    if df.empty:
        st.info("No data matches the selected filters.")
        return

    # Only dealers with a next slab (exclude Slab E)
    upgradable = df[df["Qualified Slab"] != SLAB_CONFIG[-1]["slab"]].copy()

    if upgradable.empty:
        st.info("All dealers are already at the highest slab.")
        return

    # Compute points gap
    upgradable["Points Gap"] = upgradable.apply(
        lambda row: points_to_next(row["Qualified Points"], row["Qualified Slab"]) or 0,
        axis=1,
    )

    # --- Threshold selector ---
    threshold = st.slider(
        "Show dealers within N points of next slab",
        min_value=100,
        max_value=2000,
        value=500,
        step=100,
        key="near_upgrade_threshold",
    )

    near = upgradable[upgradable["Points Gap"] <= threshold].copy()
    near = near.sort_values("Points Gap", ascending=True)

    # --- KPI row ---
    total_near = len(near)
    st.markdown(f'<div class="section-title">Dealers within {threshold} points of upgrade</div>',
                unsafe_allow_html=True)

    if total_near == 0:
        st.info(f"No dealers are within {threshold} points of the next slab.")
        return

    # Breakdown by target slab
    near["Target Slab"] = near["Qualified Slab"].apply(get_next_slab)
    target_counts = near["Target Slab"].value_counts()

    kpi_cols = st.columns(min(len(target_counts) + 1, 6))
    with kpi_cols[0]:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">Total Near-Upgrade</div>
                <div class="kpi-value">{total_near}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    for i, (target, count) in enumerate(target_counts.items()):
        if i + 1 >= len(kpi_cols):
            break
        color = SLAB_COLORS.get(target, "#64748b")
        with kpi_cols[i + 1]:
            st.markdown(
                f"""
                <div class="kpi-card" style="border-top-color: {color};">
                    <div class="kpi-label">Near {target}</div>
                    <div class="kpi-value" style="color: {color};">{count}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # --- Detail table ---
    st.subheader(f"Near-Upgrade Dealers ({total_near} records)")

    display_cols = [
        c for c in [
            "Dealer Name", "Distributor Name", "State", "Zone",
            "Qualified Points", "Qualified Slab", "Target Slab", "Points Gap",
            "Total Volume",
        ]
        if c in near.columns
    ]

    display_df = near[display_cols].copy()

    # Rename for display
    display_df = display_df.rename(columns={
        "Target Slab": "Next Slab",
        "Points Gap": "Pts to Upgrade",
    })

    # Round numerics to whole numbers
    for col in ["Qualified Points", "Pts to Upgrade", "Total Volume"]:
        if col in display_df.columns:
            display_df[col] = pd.to_numeric(
                display_df[col], errors="coerce"
            ).fillna(0).round(0).astype(int)

    def _highlight_urgency(row: pd.Series) -> list[str]:
        """Color rows by upgrade urgency: green for close, yellow for moderate."""
        pts = row.get("Pts to Upgrade", 999)
        if pts <= 100:
            return ["background-color: #d1fae5"] * len(row)
        if pts <= 300:
            return ["background-color: #fef3c7"] * len(row)
        return ["background-color: #f1f5f9"] * len(row)

    def _bold_key_cols(col: pd.Series) -> list[str]:
        """Bold key columns."""
        if col.name in ("Dealer Name", "Pts to Upgrade", "Next Slab"):
            return ["font-weight: 700"] * len(col)
        return [""] * len(col)

    styled = (
        display_df.style
        .apply(_highlight_urgency, axis=1)
        .apply(_bold_key_cols, axis=0)
    )
    st.dataframe(styled, use_container_width=True, hide_index=True, height=500)


# ============================================================================
# SECTION 9 — TAB: PERFORMANCE OVERVIEW (Distributor / Zone / State)
# ============================================================================

def _build_performance_table(
    df: pd.DataFrame,
    group_col: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Aggregate dealer metrics by a grouping column with slab distribution.

    Args:
        df: Source DataFrame.
        group_col: Column to group by (e.g. "Distributor Name", "Zone", "State").

    Returns:
        Tuple of (aggregated DataFrame with Qual. Rate %, slab_mix DataFrame).
    """
    agg = df.groupby(group_col, as_index=False).agg(
        Dealers=("Dealer Name", "count"),
        Total_Volume=("Total Volume", "sum"),
        Avg_Points=("Qualified Points", "mean"),
        Total_Points=("Qualified Points", "sum"),
        Shop_Volume=("Shop Volume", "sum"),
        Site_Volume=("Total Site Volume", "sum"),
    )
    agg = agg[agg[group_col].str.strip() != ""]
    agg = agg.sort_values("Dealers", ascending=False)

    # Slab mix
    valid = df[df[group_col].str.strip() != ""]
    slab_mix = valid.groupby([group_col, "Qualified Slab"]).size().unstack(fill_value=0)
    for slab_name in SLAB_ORDER:
        if slab_name not in slab_mix.columns:
            slab_mix[slab_name] = 0
    slab_mix = slab_mix[SLAB_ORDER]

    if "Unqualified" in slab_mix.columns:
        slab_mix["Qualified Rate"] = (
            slab_mix.drop(columns=["Unqualified"]).sum(axis=1)
            / slab_mix.sum(axis=1)
            * 100
        )
    else:
        slab_mix["Qualified Rate"] = 100.0

    agg = agg.merge(
        slab_mix[["Qualified Rate"]],
        left_on=group_col,
        right_index=True,
        how="left",
    )

    return agg, slab_mix


def _render_performance_section(
    agg_df: pd.DataFrame,
    slab_mix: pd.DataFrame,
    group_col: str,
    title: str,
    height: int = 400,
) -> None:
    """Render a styled performance table for a given grouping.

    Args:
        agg_df: Aggregated metrics DataFrame.
        slab_mix: Slab distribution DataFrame.
        group_col: Grouping column name.
        title: Subheader title text.
        height: Table height in pixels.
    """
    total_rows = len(agg_df)
    st.subheader(f"{title} ({total_rows})")

    display_df = agg_df.copy()
    display_df = display_df.rename(columns={
        "Total_Volume": "Total Volume",
        "Avg_Points": "Avg Points",
        "Total_Points": "Total Points",
        "Shop_Volume": "Qual. Shop Vol.",
        "Site_Volume": "Qual. Site Vol.",
        "Qualified Rate": "Qual. Rate %",
    })

    # Add slab distribution columns
    for slab_name in SLAB_ORDER:
        if slab_name in slab_mix.columns:
            display_df = display_df.merge(
                slab_mix[[slab_name]],
                left_on=group_col,
                right_index=True,
                how="left",
            )

    display_cols = [
        c for c in [
            group_col, "Dealers", "Total Volume",
            "Qual. Shop Vol.", "Qual. Site Vol.",
            "Avg Points", "Total Points", "Qual. Rate %",
        ] + SLAB_ORDER
        if c in display_df.columns
    ]
    display_df = display_df[display_cols]

    # All numerics to whole numbers
    int_cols = ["Dealers", "Total Volume", "Qual. Shop Vol.", "Qual. Site Vol.",
                "Total Points", "Avg Points", "Qual. Rate %"] + SLAB_ORDER
    for col in int_cols:
        if col in display_df.columns:
            display_df[col] = pd.to_numeric(
                display_df[col], errors="coerce"
            ).fillna(0).round(0).astype(int)

    def _highlight_qual(row: pd.Series) -> list[str]:
        rate = row.get("Qual. Rate %", 0)
        if rate >= 60:
            return ["background-color: #d1fae5"] * len(row)
        if rate >= 30:
            return ["background-color: #fef3c7"] * len(row)
        return ["background-color: #fee2e2"] * len(row)

    def _bold_key(col_series: pd.Series) -> list[str]:
        if col_series.name in (group_col, "Dealers", "Qual. Rate %"):
            return ["font-weight: 700"] * len(col_series)
        return [""] * len(col_series)

    styled = (
        display_df.style
        .apply(_highlight_qual, axis=1)
        .apply(_bold_key, axis=0)
    )
    st.dataframe(styled, use_container_width=True, hide_index=True, height=height)


def render_performance_overview(df: pd.DataFrame) -> None:
    """Render the Performance Overview tab with Distributor, Zone, and State tables.

    Args:
        df: Filtered DataFrame.
    """
    if df.empty:
        st.info("No data available for performance analysis.")
        return

    # --- KPI cards ---
    st.markdown('<div class="section-title">Performance Overview</div>',
                unsafe_allow_html=True)

    dist_count = df["Distributor Name"].nunique() if "Distributor Name" in df.columns else 0
    zone_count = df["Zone"][df["Zone"].str.strip() != ""].nunique() if "Zone" in df.columns else 0
    state_count = df["State"][df["State"].str.strip() != ""].nunique() if "State" in df.columns else 0
    total_dealers = len(df)

    kpi_cols = st.columns(4)
    kpis = [
        ("Total Dealers", format_indian(total_dealers)),
        ("Distributors", format_indian(dist_count)),
        ("Zones", format_indian(zone_count)),
        ("States", format_indian(state_count)),
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

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # --- Zone-wise table ---
    if "Zone" in df.columns:
        zone_agg, zone_slab = _build_performance_table(df, "Zone")
        _render_performance_section(zone_agg, zone_slab, "Zone", "Zone-wise Performance", height=200)
        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # --- State-wise table ---
    if "State" in df.columns:
        state_agg, state_slab = _build_performance_table(df, "State")
        _render_performance_section(state_agg, state_slab, "State", "State-wise Performance")
        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # --- Distributor-wise table ---
    if "Distributor Name" in df.columns:
        dist_agg, dist_slab = _build_performance_table(df, "Distributor Name")
        _render_performance_section(dist_agg, dist_slab, "Distributor Name", "Distributor-wise Performance")


# ============================================================================
# SECTION 10 — TAB: COSTING ANALYSIS
# ============================================================================

def render_costing(df: pd.DataFrame) -> None:
    """Render the Costing Analysis tab with gift cost and per-MT breakdown.

    Gift cost formula: Gifts = Total Points in Slab / Slab Threshold Points,
    then Total Cost = Gifts x Gift Value (INR).

    Args:
        df: Filtered DataFrame.
    """
    if df.empty:
        st.info("No data available for costing analysis.")
        return

    # --- Compute costing per slab ---
    costing_rows = []
    grand_dealers = 0
    grand_points = 0.0
    grand_gifts = 0
    grand_cost = 0.0

    for cfg in SLAB_CONFIG:
        slab_df = df[df["Qualified Slab"] == cfg["slab"]]
        count = len(slab_df)
        total_pts = slab_df["Qualified Points"].sum() if "Qualified Points" in slab_df.columns else 0
        threshold = cfg["threshold_points"]
        gift_inr = cfg["gift_inr"]

        if threshold > 0 and total_pts > 0:
            gifts = int(total_pts / threshold)
            cost = gifts * gift_inr
        else:
            gifts = 0
            cost = 0

        grand_dealers += count
        grand_points += total_pts
        grand_gifts += gifts
        grand_cost += cost

        costing_rows.append({
            "Slab": cfg["slab"],
            "Points Range": cfg["range"],
            "Dealers": count,
            "Total Points": format_indian(total_pts),
            "Threshold Pts": format_indian(threshold) if threshold > 0 else "-",
            "Gifts": format_indian(gifts),
            "Gift Value (INR)": format_indian(gift_inr, prefix="₹") if gift_inr > 0 else "-",
            "Total Cost (INR)": format_indian(cost, prefix="₹"),
            "% of Total": f"{cost / grand_cost * 100:.1f}%" if grand_cost > 0 else "0%",
        })

    # Recalculate % after grand total is known
    for row in costing_rows:
        cost_str = row["Total Cost (INR)"]
        # Find actual cost from the slab
        slab_name = row["Slab"]
        cfg = next(c for c in SLAB_CONFIG if c["slab"] == slab_name)
        slab_df = df[df["Qualified Slab"] == slab_name]
        total_pts = slab_df["Qualified Points"].sum() if "Qualified Points" in slab_df.columns else 0
        threshold = cfg["threshold_points"]
        gift_inr = cfg["gift_inr"]
        if threshold > 0 and total_pts > 0:
            cost = int(total_pts / threshold) * gift_inr
        else:
            cost = 0
        row["% of Total"] = f"{cost / grand_cost * 100:.1f}%" if grand_cost > 0 else "0%"

    # --- Editable Retail Sales input ---
    retail_sales = st.number_input(
        "Total Retail Sales (MT)",
        min_value=0.0,
        value=TOTAL_RETAIL_SALES,
        step=100.0,
        format="%.0f",
        key="costing_retail_sales",
    )

    per_mt = grand_cost / retail_sales if retail_sales > 0 else 0

    # --- KPI cards ---
    st.markdown('<div class="section-title">Costing Overview</div>', unsafe_allow_html=True)

    kpi_cols = st.columns(4)
    kpis = [
        ("Total Gifting Cost", format_indian(grand_cost, prefix="₹")),
        ("Cost per MT", f"₹{per_mt:,.2f}"),
        ("Total Gifts", format_indian(grand_gifts)),
        ("Total Retail Sales (MT)", format_indian(retail_sales)),
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

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # --- TOTAL row ---
    costing_rows.append({
        "Slab": "TOTAL",
        "Points Range": "",
        "Dealers": grand_dealers,
        "Total Points": format_indian(grand_points),
        "Threshold Pts": "",
        "Gifts": format_indian(grand_gifts),
        "Gift Value (INR)": "",
        "Total Cost (INR)": format_indian(grand_cost, prefix="₹"),
        "% of Total": "100%",
    })

    # --- Table ---
    st.subheader("Slab-wise Costing Breakdown")
    costing_df = pd.DataFrame(costing_rows)

    def _style_costing_row(row: pd.Series) -> list[str]:
        """Apply styling to costing table rows."""
        slab = row.get("Slab", "")
        if slab == "TOTAL":
            return ["font-weight: 800; background-color: #cbd5e1; color: #0f172a"] * len(row)
        bg = SLAB_COLORS_LIGHT.get(slab, "")
        if bg:
            return [f"background-color: {bg}"] * len(row)
        return [""] * len(row)

    def _bold_costing_cols(col: pd.Series) -> list[str]:
        """Bold key costing columns."""
        if col.name in ("Slab", "Gifts", "Total Cost (INR)", "% of Total"):
            return ["font-weight: 700"] * len(col)
        return [""] * len(col)

    styled = (
        costing_df.style
        .apply(_style_costing_row, axis=1)
        .apply(_bold_costing_cols, axis=0)
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)


# ============================================================================
# SECTION 11 — MAIN
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
            <h1>Q4 Scheme Dashboard</h1>
            <div class="header-right">
                <span class="header-subtitle">Slab Analysis & Dealer Tracker</span>
                <span class="header-badge">Q4 FY 26</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --- Load Data ---
    df = load_data()

    # --- File info ---
    file_path = _find_excel_file()
    st.caption(f"Data source: `{file_path.name}` — {len(df)} dealers loaded (excl. self-counter)")

    # --- Global Cascading Filters ---
    st.markdown(
        '<div class="filter-panel"><div class="filter-panel-title">&#9660; Filters</div>',
        unsafe_allow_html=True,
    )
    summary_filters = ["Qualified Slab", "Zone", "State", "District", "Distributor Name"]
    filtered_df = render_cascading_filters(df, key="global", filter_fields=summary_filters)
    st.markdown('</div>', unsafe_allow_html=True)

    # --- Summary on top (filtered) ---
    render_summary_top(filtered_df)

    # --- Sidebar PIN for Costing ---
    with st.sidebar:
        st.markdown("**Admin Access**")
        pin_input = st.text_input("Enter PIN to unlock Costing", type="password", key="costing_pin")
        show_costing = pin_input == "0000"
        if pin_input and not show_costing:
            st.error("Incorrect PIN")

    # --- Tabs ---
    if show_costing:
        tab_summary, tab_details, tab_upgrade, tab_performance, tab_costing = st.tabs([
            "📊 Summary",
            "🔍 Dealer Details",
            "🎯 Near-Upgrade",
            "📈 Performance Overview",
            "💰 Costing",
        ])
    else:
        tab_summary, tab_details, tab_upgrade, tab_performance = st.tabs([
            "📊 Summary",
            "🔍 Dealer Details",
            "🎯 Near-Upgrade",
            "📈 Performance Overview",
        ])

    with tab_summary:
        render_summary(filtered_df)

    with tab_details:
        render_dealer_details(filtered_df)

    with tab_upgrade:
        render_near_upgrade(filtered_df)

    with tab_performance:
        render_performance_overview(filtered_df)

    if show_costing:
        with tab_costing:
            render_costing(filtered_df)


if __name__ == "__main__":
    main()
