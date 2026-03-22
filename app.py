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
        "volume_mt": 0,
        "color": "#94a3b8",
        "color_light": "#f1f5f9",
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
    },
]

# Derived lookup dictionaries — all computed from SLAB_CONFIG
SLAB_GIFT_MAP: dict[str, str] = {s["slab"]: s["gift_full"] for s in SLAB_CONFIG}
SLAB_COLORS: dict[str, str] = {s["slab"]: s["color"] for s in SLAB_CONFIG}
SLAB_COLORS_LIGHT: dict[str, str] = {s["slab"]: s["color_light"] for s in SLAB_CONFIG}
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
            font-size: 0.9rem;
            font-weight: 500;
            color: #cbd5e1;
        }

        /* Slab cards */
        .slab-card {
            background: #ffffff;
            border-radius: 0.6rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.12);
            padding: 1.1rem 1.3rem;
            border-left: 5px solid #bfdbfe;
            margin-bottom: 0.5rem;
        }
        .slab-count {
            font-size: 2rem;
            font-weight: 800;
            color: #0f172a;
            line-height: 1.2;
        }
        .slab-label {
            font-size: 0.85rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #1e293b;
            margin-top: 0.3rem;
        }
        .slab-gift {
            font-size: 0.78rem;
            font-weight: 500;
            color: #475569;
            margin-top: 0.2rem;
        }

        /* KPI cards */
        .kpi-card {
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 0.6rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.10);
            padding: 1.1rem;
            text-align: center;
        }
        .kpi-label {
            font-size: 0.82rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #334155;
            margin-bottom: 0.3rem;
        }
        .kpi-value {
            font-size: 1.5rem;
            font-weight: 800;
            color: #0f172a;
        }

        /* Total row styling for tables */
        .total-row {
            font-weight: 800;
            background-color: #e2e8f0;
        }

        /* Table styling */
        .stDataFrame [data-testid="stDataFrameResizable"] {
            border: 1px solid #94a3b8;
            border-radius: 0.5rem;
        }
        .stDataFrame thead tr th {
            background-color: #1e293b !important;
            color: #ffffff !important;
            font-weight: 800 !important;
            font-size: 0.85rem !important;
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
            background-color: #f1f5f9 !important;
        }
        .stDataFrame tbody tr:nth-child(odd) {
            background-color: #ffffff !important;
        }
        .stDataFrame tbody tr:hover {
            background-color: #e2e8f0 !important;
        }

        /* Hide Streamlit chrome */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        /* Streamlit tab styling */
        .stTabs [data-baseweb="tab-list"] {
            gap: 0.5rem;
            border-bottom: 2px solid #cbd5e1;
        }
        .stTabs [data-baseweb="tab"] {
            padding: 0.6rem 1.2rem;
            font-size: 0.9rem;
            font-weight: 600;
            color: #334155;
        }
        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            font-weight: 700;
            color: #0f172a;
            border-bottom: 3px solid #2563eb;
        }

        /* Filter dropdowns */
        .stSelectbox label {
            font-weight: 600 !important;
            color: #1e293b !important;
            font-size: 0.85rem !important;
        }
        .stSelectbox [data-baseweb="select"] {
            border-color: #94a3b8 !important;
        }

        /* Subheaders */
        .stSubheader, h3, h2 {
            color: #0f172a !important;
            font-weight: 700 !important;
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
    """Render the top summary section with KPIs and slab cards.

    Args:
        df: Filtered DataFrame.
    """
    # --- KPI Row ---
    total_dealers = len(df)
    total_volume = df["Total Volume"].sum() if "Total Volume" in df.columns else 0
    shop_qual_vol = df["Shop Volume"].sum() if "Shop Volume" in df.columns else 0
    site_qual_vol = df["Site Volume"].sum() if "Site Volume" in df.columns else 0
    total_points = df["Qualified Points"].sum() if "Qualified Points" in df.columns else 0

    total_shop_vol = df["Shop Volume"].sum() if "Shop Volume" in df.columns else 0
    total_site_vol = df["Site Volume"].sum() if "Site Volume" in df.columns else 0
    total_qual_vol = df["Qualified Volume"].sum() if "Qualified Volume" in df.columns else 0

    kpi_cols = st.columns(6)
    kpis = [
        ("Total Dealers", format_indian(total_dealers)),
        ("Total Volume (MT)", format_indian(total_volume, decimal=1)),
        ("Qual. Shop Vol. (MT)", format_indian(total_shop_vol, decimal=1)),
        ("Qual. Site Vol. (MT)", format_indian(total_site_vol, decimal=1)),
        ("Qualified Volume (MT)", format_indian(total_qual_vol, decimal=1)),
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
        site_vol = slab_df["Site Volume"].sum() if "Site Volume" in slab_df.columns else 0
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
            "Shop Volume", "Site Volume", "Total Volume",
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
        "Site Volume": "Qual. Site Vol.",
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

    # --- Global Cascading Filters ---
    summary_filters = ["Qualified Slab", "Zone", "State", "District", "Distributor Name"]
    filtered_df = render_cascading_filters(df, key="global", filter_fields=summary_filters)

    # --- Summary on top (filtered) ---
    render_summary_top(filtered_df)

    # --- Tabs ---
    tab_summary, tab_details = st.tabs([
        "📊 Summary",
        "🔍 Dealer Details",
    ])

    with tab_summary:
        render_summary(filtered_df)

    with tab_details:
        render_dealer_details(filtered_df)


if __name__ == "__main__":
    main()
