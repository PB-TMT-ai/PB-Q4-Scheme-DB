"""
Streamlit Dashboard — Q4 Scheme Slab Analysis (V2 — Bento Grid)
================================================================
Bento Grid redesign of the Q4 Scheme Dashboard.
Apple-style modular cards: mixed sizes, rounded corners, soft shadows,
Fira Sans/Code fonts, #F5F5F7 background, hover scale effects.
Same data pipeline as app.py.

Run: streamlit run app_v2.py
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
DATA_FILE: str = "Q4 as on 31st Mar.xlsx"
SHEET_NAME: str = "Sheet1"
HEADER_ROW: int = 0

COLUMN_MAP: dict[str, str] = {
    "Retailer Name": "Dealer Name",
    "State Name": "State",
    "District Name": "District",
    "Shop vol.": "Shop Volume",
    "Site vol.": "Site Volume",
    "Total vol. under scheme": "Qualified Volume",
    "Q4 Vol. under eligible scheme": "Qualified Volume",
    "Total site vol.": "Total Site Volume",
    "Site vol. under scheme": "Total Site Volume",
    "Points": "Qualified Points",
    "Current gift": "Gift",
    "Current gift slab": "Current Slab",
    "Distributor self-counter (Yes/No)": "Self Counter",
    "Jan + Feb+Mar": "Q4 Volume",
    "# Unique Site >200 MT": "Unique Site >200 MT",
}


def _find_excel_file() -> Path:
    """Locate the configured Excel data file in DATA_DIR.

    Priority: explicit DATA_FILE → newest 'Q4 as on*.xlsx' by name → newest .xlsx by name.
    """
    all_xlsx = sorted(glob.glob(str(Path(DATA_DIR) / "*.xlsx")))
    log_info(f"Excel files in {DATA_DIR}/: {[Path(f).name for f in all_xlsx]}")

    explicit = Path(DATA_DIR) / DATA_FILE
    if explicit.exists():
        log_info(f"Using explicit data file: {explicit}")
        return explicit

    q4_files = sorted(
        glob.glob(str(Path(DATA_DIR) / "Q4 data as on*.xlsx"))
        + glob.glob(str(Path(DATA_DIR) / "Q4 as on*.xlsx")),
        reverse=True,
    )
    if q4_files:
        log_info(f"Fallback: using newest Q4 file by name: {q4_files[0]}")
        return Path(q4_files[0])

    if not all_xlsx:
        st.error(f"No .xlsx files found in `{DATA_DIR}/`. Please add your data file.")
        log_error(f"No Excel files in {DATA_DIR}/")
        st.stop()
    log_info(f"Fallback: using {all_xlsx[-1]}")
    return Path(all_xlsx[-1])


# ---------------------------------------------------------------------------
# SLAB_CONFIG — Single source of truth for all tier definitions
# ---------------------------------------------------------------------------

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

TOTAL_RETAIL_SALES: float = 47180.0

SLAB_CODE_MAP: dict[str, str] = {s["slab_code"]: s["slab"] for s in SLAB_CONFIG}

NEXT_SLAB_MAP: dict[str, Optional[str]] = {}
for _i, _cfg in enumerate(SLAB_CONFIG):
    NEXT_SLAB_MAP[_cfg["slab"]] = (
        SLAB_CONFIG[_i + 1]["slab"] if _i + 1 < len(SLAB_CONFIG) else None
    )

NEXT_SLAB_THRESHOLD: dict[str, Optional[float]] = {}
for _i, _cfg in enumerate(SLAB_CONFIG):
    NEXT_SLAB_THRESHOLD[_cfg["slab"]] = (
        SLAB_CONFIG[_i + 1]["lower"] if _i + 1 < len(SLAB_CONFIG) else None
    )

NEXT_SLAB_VOLUME: dict[str, Optional[float]] = {}
for _i, _cfg in enumerate(SLAB_CONFIG):
    NEXT_SLAB_VOLUME[_cfg["slab"]] = (
        SLAB_CONFIG[_i + 1]["volume_mt"] if _i + 1 < len(SLAB_CONFIG) else None
    )


# ============================================================================
# SECTION 2 — HELPER FUNCTIONS
# ============================================================================

def format_indian(number: float, prefix: str = "", decimal: int = 0) -> str:
    """Format a number with Indian comma grouping (e.g. 12,34,567)."""
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


def assign_slab(points: float) -> str:
    """Assign a slab label based on qualified points using SLAB_CONFIG."""
    if pd.isna(points):
        points = 0.0
    points = float(points)
    for cfg in SLAB_CONFIG:
        if cfg["lower"] <= points < cfg["upper"]:
            return cfg["slab"]
    return SLAB_CONFIG[0]["slab"]


def get_next_slab(current: str) -> Optional[str]:
    """Look up the next higher slab from NEXT_SLAB_MAP."""
    return NEXT_SLAB_MAP.get(current)


def points_to_next(points: float, current: str) -> Optional[float]:
    """Calculate the points gap to reach the next slab tier."""
    threshold = NEXT_SLAB_THRESHOLD.get(current)
    if threshold is None:
        return None
    gap = threshold - float(points)
    return max(gap, 0.0)


# ============================================================================
# SECTION 3 — BENTO GRID CSS
# ============================================================================

def inject_bento_css() -> None:
    """Inject Bento Grid (Apple-style) CSS into the Streamlit app."""
    st.markdown(
        """
        <style>
        /* ── Google Fonts: Fira Sans + Fira Code ── */
        @import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500;600;700&family=Fira+Sans:wght@300;400;500;600;700&display=swap');

        /* ── Global Reset ── */
        html, body, [class*="css"] {
            font-family: 'Fira Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        }
        .stApp {
            background-color: #F5F5F7;
        }

        /* ── Header ── */
        .v2-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 1.25rem 1.5rem;
            background: #FFFFFF;
            border-radius: 20px;
            margin-bottom: 1.25rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        }
        .v2-header h1 {
            margin: 0;
            font-size: 1.35rem;
            font-weight: 700;
            color: #1D1D1F;
            letter-spacing: -0.02em;
        }
        .v2-header-right {
            display: flex;
            align-items: center;
            gap: 1rem;
        }
        .v2-header-subtitle {
            font-size: 0.8rem;
            font-weight: 400;
            color: #86868B;
        }
        .v2-header-badge {
            font-size: 0.72rem;
            font-weight: 600;
            color: #FFFFFF;
            background: #1D1D1F;
            border-radius: 100px;
            padding: 0.3rem 0.9rem;
            letter-spacing: 0.04em;
        }

        /* ── Bento Card Base ── */
        .v2-bento-card {
            background: #FFFFFF;
            border-radius: 20px;
            padding: 1.5rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.04);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .v2-bento-card:hover {
            transform: scale(1.01);
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }

        /* ── KPI Metrics (inside bento cards) ── */
        .v2-kpi {
            text-align: left;
            padding: 0;
        }
        .v2-kpi-value {
            font-family: 'Fira Code', monospace;
            font-size: 2rem;
            font-weight: 700;
            color: #1D1D1F;
            line-height: 1.1;
            letter-spacing: -0.03em;
        }
        .v2-kpi-label {
            font-size: 0.72rem;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: #86868B;
            margin-top: 0.5rem;
        }

        /* ── Slab Cards (bento tiles) ── */
        .v2-slab {
            background: #FFFFFF;
            border-radius: 20px;
            padding: 1.25rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.04);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
            cursor: default;
            position: relative;
            overflow: hidden;
        }
        .v2-slab:hover {
            transform: scale(1.02);
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }
        .v2-slab-accent {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            border-radius: 20px 20px 0 0;
        }
        .v2-slab-count {
            font-family: 'Fira Code', monospace;
            font-size: 2rem;
            font-weight: 700;
            color: #1D1D1F;
            line-height: 1.1;
            letter-spacing: -0.03em;
            margin-top: 0.5rem;
        }
        .v2-slab-name {
            font-size: 0.82rem;
            font-weight: 600;
            color: #1D1D1F;
            margin-top: 0.4rem;
        }
        .v2-slab-range {
            font-size: 0.7rem;
            font-weight: 400;
            color: #86868B;
            margin-top: 0.15rem;
        }
        .v2-slab-gift {
            font-size: 0.7rem;
            font-weight: 500;
            color: #86868B;
            margin-top: 0.3rem;
            padding-top: 0.3rem;
            border-top: 1px solid #F5F5F7;
        }

        /* ── Section Headers ── */
        .v2-section-title {
            font-size: 0.82rem;
            font-weight: 600;
            color: #1D1D1F;
            text-transform: uppercase;
            letter-spacing: 0.1em;
            margin: 0.5rem 0 1rem 0;
            padding: 0;
        }

        /* ── Tables inside bento cards ── */
        .stDataFrame [data-testid="stDataFrameResizable"] {
            border: none;
            border-radius: 12px;
            overflow: hidden;
        }
        .stDataFrame thead tr th {
            background-color: #F5F5F7 !important;
            color: #1D1D1F !important;
            font-weight: 600 !important;
            font-size: 0.78rem !important;
            border-bottom: none !important;
        }
        [data-testid="stDataFrame"] [role="columnheader"],
        [data-testid="stDataFrame"] [data-testid="glide-cell"] {
            font-weight: 600 !important;
        }
        [data-testid="stDataFrame"] .gdg-header {
            font-weight: 600 !important;
        }
        .stDataFrame tbody tr {
            background-color: #FFFFFF !important;
        }
        .stDataFrame tbody tr:hover {
            background-color: #F9F9FB !important;
        }

        /* ── Tabs — Pill style ── */
        .stTabs [data-baseweb="tab-list"] {
            gap: 0.5rem;
            border-bottom: none;
            background: #FFFFFF;
            border-radius: 14px;
            padding: 0.3rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        }
        .stTabs [data-baseweb="tab"] {
            padding: 0.5rem 1.25rem;
            font-size: 0.8rem;
            font-weight: 500;
            color: #86868B;
            border-radius: 10px;
            border-bottom: none;
        }
        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            font-weight: 600;
            color: #1D1D1F;
            background: #F5F5F7;
            border-bottom: none;
        }

        /* ── Filters ── */
        .stSelectbox label {
            font-weight: 500 !important;
            color: #1D1D1F !important;
            font-size: 0.76rem !important;
            letter-spacing: 0.02em !important;
        }
        .stSelectbox [data-baseweb="select"] {
            border-color: #E8E8ED !important;
            border-radius: 12px !important;
        }

        /* ── Subheaders ── */
        .stSubheader, h3, h2 {
            color: #1D1D1F !important;
            font-weight: 600 !important;
        }

        /* ── File caption ── */
        .v2-caption {
            font-size: 0.7rem;
            color: #86868B;
            margin-bottom: 1rem;
        }

        /* ── Hide Streamlit chrome ── */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================================
# SECTION 4 — DATA LOADING
# ============================================================================

@st.cache_data(ttl=3600)
def load_data() -> pd.DataFrame:
    """Load, clean, and return the Excel data."""
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

    rename_map = {k: v for k, v in COLUMN_MAP.items() if k in df.columns}
    df = df.rename(columns=rename_map)
    log_info(f"Renamed columns: {list(rename_map.values())}")

    numeric_cols = [
        "Qualified Volume", "Total Site Volume", "Qualified Points",
        "Shop Volume", "Site Volume", "Q4 Volume", "Unique Site >200 MT",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

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

    if "Qualified Points" in df.columns:
        df["Qualified Slab"] = df["Qualified Points"].apply(assign_slab)
    elif "Current Slab" in df.columns:
        df["Current Slab"] = df["Current Slab"].fillna("-").astype(str).str.strip()
        df["Qualified Slab"] = df["Current Slab"].map(SLAB_CODE_MAP).fillna("Unqualified")
    else:
        df["Qualified Slab"] = "Unqualified"

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
    """Build filter options: ['All'] + sorted unique non-blank values."""
    unique_vals = series[series.astype(str).str.strip() != ""].unique()
    return ["All"] + sorted(str(v) for v in unique_vals if str(v).strip())


def render_cascading_filters(
    df: pd.DataFrame,
    key: str,
    filter_fields: Optional[list[str]] = None,
) -> pd.DataFrame:
    """Render cascading dropdown filters and return the filtered DataFrame."""
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

def render_kpi_row(df: pd.DataFrame) -> None:
    """Render the top KPI metrics as bento grid cards."""
    total_dealers = len(df)
    total_volume = df["Total Volume"].sum() if "Total Volume" in df.columns else 0
    total_shop_vol = df["Shop Volume"].sum() if "Shop Volume" in df.columns else 0
    total_site_vol = df["Total Site Volume"].sum() if "Total Site Volume" in df.columns else 0
    _qual_mask = df["Qualified Slab"] != "Unqualified" if "Qualified Slab" in df.columns else pd.Series([True] * len(df))
    total_qual_vol = df.loc[_qual_mask, "Qualified Volume"].sum() if "Qualified Volume" in df.columns else 0
    total_points = df["Qualified Points"].sum() if "Qualified Points" in df.columns else 0

    # Row 1: 2-wide hero card + 2 standard cards
    r1c1, r1c2, r1c3 = st.columns([2, 1, 1])
    with r1c1:
        st.markdown(
            f"""
            <div class="v2-bento-card">
                <div class="v2-kpi">
                    <div class="v2-kpi-value">{format_indian(total_dealers)}</div>
                    <div class="v2-kpi-label">Total Dealers</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with r1c2:
        st.markdown(
            f"""
            <div class="v2-bento-card">
                <div class="v2-kpi">
                    <div class="v2-kpi-value">{format_indian(total_volume, decimal=1)}</div>
                    <div class="v2-kpi-label">Total Volume (MT)</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with r1c3:
        st.markdown(
            f"""
            <div class="v2-bento-card">
                <div class="v2-kpi">
                    <div class="v2-kpi-value">{format_indian(total_points, decimal=1)}</div>
                    <div class="v2-kpi-label">Total Qualified Points</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown('<div style="height: 0.75rem;"></div>', unsafe_allow_html=True)

    # Row 2: 3 equal cards
    r2c1, r2c2, r2c3 = st.columns(3)
    with r2c1:
        st.markdown(
            f"""
            <div class="v2-bento-card">
                <div class="v2-kpi">
                    <div class="v2-kpi-value">{format_indian(total_shop_vol, decimal=1)}</div>
                    <div class="v2-kpi-label">Qual. Shop Vol. (MT)</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with r2c2:
        st.markdown(
            f"""
            <div class="v2-bento-card">
                <div class="v2-kpi">
                    <div class="v2-kpi-value">{format_indian(total_site_vol, decimal=1)}</div>
                    <div class="v2-kpi-label">Qual. Site Vol. (MT)</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with r2c3:
        st.markdown(
            f"""
            <div class="v2-bento-card">
                <div class="v2-kpi">
                    <div class="v2-kpi-value">{format_indian(total_qual_vol, decimal=1)}</div>
                    <div class="v2-kpi-label">Qualified Volume (MT)</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_slab_cards(df: pd.DataFrame) -> None:
    """Render slab distribution as bento grid tiles with colored accent bars."""
    if "Qualified Slab" not in df.columns:
        return

    slab_counts = df["Qualified Slab"].value_counts()
    card_cols = st.columns(len(SLAB_CONFIG))
    for col, cfg in zip(card_cols, SLAB_CONFIG):
        count = int(slab_counts.get(cfg["slab"], 0))
        color = cfg["color"]
        with col:
            st.markdown(
                f"""
                <div class="v2-slab">
                    <div class="v2-slab-accent" style="background-color: {color};"></div>
                    <div class="v2-slab-count">{count}</div>
                    <div class="v2-slab-name">{cfg['slab']}</div>
                    <div class="v2-slab-range">{cfg['range']} pts</div>
                    <div class="v2-slab-gift">{cfg['gift']}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def render_summary(df: pd.DataFrame) -> None:
    """Render the Summary tab with slab breakdown table."""
    st.markdown('<div class="v2-section-title">Slab Breakdown</div>', unsafe_allow_html=True)

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
            "Total Volume": format_indian(vol, decimal=1),
            "Qual. Shop Vol.": format_indian(shop_vol, decimal=1),
            "Qual. Site Vol.": format_indian(site_vol, decimal=1),
            "Qualified Volume": format_indian(qual_vol, decimal=1),
            "Total Points": format_indian(pts, decimal=1),
            "Gift": cfg["gift_full"],
        })

    summary_rows.append({
        "Slab": "TOTAL",
        "Points Range": "",
        "Dealer Count": grand_count,
        "Total Volume": format_indian(grand_vol, decimal=1),
        "Qual. Shop Vol.": format_indian(grand_shop_vol, decimal=1),
        "Qual. Site Vol.": format_indian(grand_site_vol, decimal=1),
        "Qualified Volume": format_indian(grand_qual_vol, decimal=1),
        "Total Points": format_indian(grand_pts, decimal=1),
        "Gift": "",
    })

    summary_df = pd.DataFrame(summary_rows)

    def _style_summary_row(row: pd.Series) -> list[str]:
        """Apply minimal styling to summary rows."""
        slab = row.get("Slab", "")
        if slab == "TOTAL":
            return ["font-weight: 700; background-color: #F3F4F6; color: #111111"] * len(row)
        return [""] * len(row)

    def _bold_summary_columns(col: pd.Series) -> list[str]:
        """Bold key summary columns."""
        if col.name in ("Slab", "Dealer Count", "Total Points"):
            return ["font-weight: 600"] * len(col)
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
    """Calculate volume remaining to reach the next slab tier."""
    next_vol = NEXT_SLAB_VOLUME.get(slab)
    if next_vol is None:
        return 0
    vol = float(total_volume) if not pd.isna(total_volume) else 0.0
    gap = next_vol - vol
    return max(gap, 0)


def render_dealer_details(df: pd.DataFrame) -> None:
    """Render the Dealer Details tab with per-dealer table and filters."""
    detail_filters = ["Distributor Name", "State", "Dealer Name", "Qualified Slab"]
    _all_cols = st.columns(len(detail_filters) + 1)
    filtered = df.copy()

    for i, field in enumerate(detail_filters):
        with _all_cols[i]:
            options = _opts(filtered[field])
            selected = st.selectbox(
                field,
                options,
                key=f"v2_dealer_detail_{field}",
            )
            if selected != "All":
                filtered = filtered[filtered[field] == selected]

    _vol_options = ["All", "0 - 20 MT", "21 - 40 MT", "41 - 60 MT", "61 - 100 MT", "More than 100 MT"]
    with _all_cols[-1]:
        _vol_sel = st.selectbox("Vol. to Achieve", _vol_options, key="v2_dealer_detail_vol_achieve")

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

    source_cols = [
        c for c in [
            "Dealer Name", "Distributor Name", "State", "Zone",
            "Shop Volume", "Total Site Volume", "Total Volume",
            "Qualified Slab", "Next Upgrade Slab",
            "Qualified Points",
        ]
        if c in filtered.columns
    ]

    st.markdown(
        f'<div class="v2-section-title">Dealer Details — {len(filtered)} records</div>',
        unsafe_allow_html=True,
    )

    display_df = filtered[source_cols].copy()

    rename_map: dict[str, str] = {
        "Shop Volume": "Qual. Shop Vol.",
        "Total Site Volume": "Qual. Site Vol.",
        "Next Upgrade Slab": "Next Slab",
    }
    display_df = display_df.rename(columns=rename_map)

    if "Next Slab" in display_df.columns:
        display_df["Next Slab"] = display_df["Next Slab"].fillna("-")

    display_df["Vol. to Achieve"] = display_df.apply(
        lambda row: _calc_vol_to_achieve(
            row.get("Qualified Slab", ""),
            row.get("Total Volume", 0),
        ),
        axis=1,
    )

    ordered_cols = [
        c for c in [
            "Dealer Name", "Distributor Name", "State", "Zone",
            "Qual. Shop Vol.", "Qual. Site Vol.", "Total Volume",
            "Qualified Slab", "Next Slab", "Vol. to Achieve", "Qualified Points",
        ]
        if c in display_df.columns
    ]
    display_df = display_df[ordered_cols]

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
        """Apply subtle slab-based styling to each row."""
        for field in ["Dealer Name", "Distributor Name"]:
            val = row.get(field)
            if isinstance(val, str) and "total" in val.lower():
                return ["font-weight: 600; background-color: #F3F4F6"] * len(row)
        return [""] * len(row)

    def _bold_key_columns(col: pd.Series) -> list[str]:
        """Bold key columns."""
        if col.name in ("Dealer Name", "Qualified Slab", "Qualified Points",
                         "Next Slab", "Vol. to Achieve"):
            return ["font-weight: 600"] * len(col)
        return [""] * len(col)

    styled = (
        display_df.style
        .apply(_highlight_by_slab, axis=1)
        .apply(_bold_key_columns, axis=0)
    )
    st.dataframe(styled, use_container_width=True, hide_index=True, height=500)


# ============================================================================
# SECTION 8 — TAB: COSTING ANALYSIS
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
            "% of Total": "0%",
        })

    # Calculate % after grand total is known
    for row in costing_rows:
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
        key="v2_costing_retail_sales",
    )

    per_mt = grand_cost / retail_sales if retail_sales > 0 else 0

    # --- KPI cards ---
    st.markdown('<div class="v2-section-title">Costing Overview</div>', unsafe_allow_html=True)

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
                <div class="v2-bento-card">
                    <div class="v2-kpi">
                        <div class="v2-kpi-value">{value}</div>
                        <div class="v2-kpi-label">{label}</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown('<div style="height: 1rem;"></div>', unsafe_allow_html=True)

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
    st.markdown('<div class="v2-section-title">Slab-wise Costing Breakdown</div>', unsafe_allow_html=True)
    costing_df = pd.DataFrame(costing_rows)

    def _style_costing_row(row: pd.Series) -> list[str]:
        """Apply styling to costing table rows."""
        slab = row.get("Slab", "")
        if slab == "TOTAL":
            return ["font-weight: 700; background-color: #F3F4F6; color: #111111"] * len(row)
        return [""] * len(row)

    def _bold_costing_cols(col: pd.Series) -> list[str]:
        """Bold key costing columns."""
        if col.name in ("Slab", "Gifts", "Total Cost (INR)", "% of Total"):
            return ["font-weight: 600"] * len(col)
        return [""] * len(col)

    styled = (
        costing_df.style
        .apply(_style_costing_row, axis=1)
        .apply(_bold_costing_cols, axis=0)
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)


# ============================================================================
# SECTION 9 — MAIN
# ============================================================================

def main() -> None:
    """Entry point for the Minimalist V2 Streamlit dashboard."""
    st.set_page_config(
        page_title="Q4 Scheme Dashboard — V2",
        layout="wide",
        page_icon="Q4",
    )

    inject_bento_css()

    # --- Header ---
    st.markdown(
        """
        <div class="v2-header">
            <h1>Q4 Scheme Dashboard</h1>
            <div class="v2-header-right">
                <span class="v2-header-subtitle">Slab Analysis & Dealer Tracker</span>
                <span class="v2-header-badge">Q4 FY 26</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --- Load Data ---
    df = load_data()

    # --- File info ---
    file_path = _find_excel_file()
    st.markdown(
        f'<div class="v2-caption">Data source: {file_path.name} — {len(df)} dealers loaded</div>',
        unsafe_allow_html=True,
    )

    # --- Global Cascading Filters ---
    summary_filters = ["Qualified Slab", "Zone", "State", "District", "Distributor Name"]
    filtered_df = render_cascading_filters(df, key="v2_global", filter_fields=summary_filters)

    # --- KPI Bento Grid ---
    render_kpi_row(filtered_df)

    st.markdown('<div style="height: 0.75rem;"></div>', unsafe_allow_html=True)

    # --- Slab Distribution Tiles ---
    render_slab_cards(filtered_df)

    st.markdown('<div style="height: 1rem;"></div>', unsafe_allow_html=True)

    # --- Tabs ---
    tab_summary, tab_details = st.tabs([
        "Summary",
        "Dealer Details",
    ])

    with tab_summary:
        render_summary(filtered_df)

    with tab_details:
        render_dealer_details(filtered_df)

        # --- PIN unlock for Costing ---
        st.markdown('<div style="height: 1rem;"></div>', unsafe_allow_html=True)
        pin_input = st.text_input("Enter PIN to unlock Costing", type="password", key="v2_costing_pin")
        if pin_input == "4141":
            render_costing(filtered_df)
        elif pin_input:
            st.error("Incorrect PIN")


if __name__ == "__main__":
    main()
