"""
Streamlit Dashboard — Scheme Slab Analysis
==========================================
Single-file Streamlit app structured in numbered sections.
Place your Excel data file in the data/ directory. Slabs are
assigned based on Qualified Points (derived from volume).

Run: streamlit run app.py
"""

from __future__ import annotations

import glob
import os
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src.lib.logger import error as log_error
from src.lib.logger import info as log_info

# ============================================================================
# SECTION 1 — CONSTANTS
# ============================================================================

DATA_DIR: str = "data"
SHEET_NAME: str = "Sheet1"
HEADER_ROW: int = 0

# Month column labels in financial-year order (Apr → Mar)
MONTH_LABELS: list[str] = [
    "Apr", "May", "Jun", "Jul", "Aug", "Sep",
    "Oct", "Nov", "Dec", "Jan", "Feb", "Mar",
]


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
# Bounds are based on QUALIFIED POINTS (not volume).
# Fields:
#   slab        — Short label displayed in UI
#   range       — Human-readable points range string
#   lower       — Inclusive lower bound (points)
#   upper       — Exclusive upper bound (use float("inf") for the top tier)
#   gift        — Short gift/reward label
#   gift_full   — Full gift description
#   category    — Grouping category
#   color       — Hex color for UI elements

SLAB_CONFIG: list[dict] = [
    {
        "slab": "Unqualified",
        "range": "0 – 749",
        "lower": 0,
        "upper": 750,
        "gift": "No Gift",
        "gift_full": "Below minimum qualification",
        "category": "Unqualified",
        "color": "#94a3b8",
    },
    {
        "slab": "Slab A",
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
        "range": "3,000 – 4,199",
        "lower": 3000,
        "upper": 4200,
        "gift": "Sony Sound Bar",
        "gift_full": "Sony - sound bar, woofer and speakers",
        "category": "B",
        "color": "#6366f1",
    },
    {
        "slab": "Slab C",
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

# Next-slab mapping (each slab → next higher slab, top tier → None)
NEXT_SLAB_MAP: dict[str, Optional[str]] = {}
for _i, _cfg in enumerate(SLAB_CONFIG):
    NEXT_SLAB_MAP[_cfg["slab"]] = (
        SLAB_CONFIG[_i + 1]["slab"] if _i + 1 < len(SLAB_CONFIG) else None
    )

# Points threshold to reach next slab (lower bound of next slab)
NEXT_SLAB_THRESHOLD: dict[str, Optional[float]] = {}
for _i, _cfg in enumerate(SLAB_CONFIG):
    NEXT_SLAB_THRESHOLD[_cfg["slab"]] = (
        SLAB_CONFIG[_i + 1]["lower"] if _i + 1 < len(SLAB_CONFIG) else None
    )

# ---------------------------------------------------------------------------
# POINTS_CONFIG — Qualified points calculation rules
# ---------------------------------------------------------------------------
# Milestones sorted descending so we match the highest applicable bonus first.

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


def calculate_qualified_volume(shop_vol: float, site_vol: float) -> float:
    """Calculate qualified volume from shop and site volumes.

    Site volume is assumed to be pre-calculated (already capped at 200 MT
    per unique site per quarter in the source data).

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

    Rules:
        - Below min_volume_mt (30 MT) → 0 points
        - Base: 1 MT = 25 points
        - Milestone bonuses (highest applicable):
          200+ MT → +50%, 150+ MT → +30%, 100+ MT → +20%, 50+ MT → +10%

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

    # Find highest applicable milestone bonus
    bonus_pct = 0
    for milestone in POINTS_CONFIG["milestones"]:
        if vol >= milestone["threshold"]:
            bonus_pct = milestone["bonus_pct"]
            break

    total = base_points * (1 + bonus_pct / 100)
    return round(total)


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
    # Fallback to first slab
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
    """Inject custom CSS styles into the Streamlit app.

    Defines card layouts, KPI styling, slab cards, and hides default
    Streamlit chrome (main menu, footer).
    """
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
            background: #ffffff;
            border-radius: 0.5rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            padding: 1rem 1.2rem;
            border-left: 4px solid #e2e8f0;
            margin-bottom: 0.5rem;
        }
        .slab-count {
            font-size: 1.7rem;
            font-weight: 700;
            color: #0f172a;
            line-height: 1.2;
        }
        .slab-label {
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #64748b;
            margin-top: 0.25rem;
        }
        .slab-gift {
            font-size: 0.72rem;
            color: #94a3b8;
            margin-top: 0.15rem;
        }

        /* KPI cards */
        .kpi-card {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 0.5rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            padding: 1rem;
            text-align: center;
        }
        .kpi-label {
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #64748b;
            margin-bottom: 0.25rem;
        }
        .kpi-value {
            font-size: 1.35rem;
            font-weight: 700;
            color: #0f172a;
        }

        /* Total row styling for tables */
        .total-row {
            font-weight: 700;
            background-color: #f1f5f9;
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
def load_data() -> tuple[pd.DataFrame, list[str]]:
    """Load and clean the Excel data file.

    Reads the latest .xlsx from DATA_DIR, renames month columns to short
    labels, coerces numerics, cleans text fields, and derives slab columns
    based on qualified points.

    Returns:
        Tuple of (cleaned DataFrame, list of month column names).
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

    # --- Rename month columns ---
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

    log_info(f"Month columns identified: {month_cols}")

    # --- Coerce month columns to numeric ---
    for col in month_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # --- Coerce volume/aggregate columns to numeric ---
    numeric_keywords = ["total", "volume", "qty", "quantity", "amount", "value", "shop", "site"]
    for col in df.columns:
        if isinstance(col, str) and any(kw in col.lower() for kw in numeric_keywords):
            if col not in month_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # --- Clean text columns ---
    text_cols = ["State", "Zone", "District", "Distributor Name", "Region",
                 "Dealer Name"]
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

    # --- Clean flag columns (Yes/No fields) ---
    flag_candidates = [
        c for c in df.columns
        if isinstance(c, str) and (
            "flag" in c.lower()
            or "status" in c.lower()
            or "eligible" in c.lower()
            or "self" in c.lower()
        )
    ]
    for col in flag_candidates:
        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.title()
        )

    # --- Total Volume (sum of month columns if not present) ---
    if month_cols and "Total Volume" not in df.columns:
        df["Total Volume"] = df[month_cols].sum(axis=1)

    # --- Shop Volume / Site Volume ---
    # Look for columns containing "shop" and "site" for volume bifurcation
    shop_col = _find_column(df, ["Shop Volume", "Shop Vol", "Shop"])
    site_col = _find_column(df, ["Site Volume", "Site Vol", "Site"])

    if shop_col:
        df["Shop Volume"] = pd.to_numeric(df[shop_col], errors="coerce").fillna(0.0)
        if shop_col != "Shop Volume":
            log_info(f"Mapped '{shop_col}' → 'Shop Volume'")
    else:
        df["Shop Volume"] = 0.0
        log_info("No Shop Volume column found — defaulting to 0")

    if site_col:
        df["Site Volume"] = pd.to_numeric(df[site_col], errors="coerce").fillna(0.0)
        if site_col != "Site Volume":
            log_info(f"Mapped '{site_col}' → 'Site Volume'")
    else:
        df["Site Volume"] = 0.0
        log_info("No Site Volume column found — defaulting to 0")

    # --- Derived columns ---
    df["Qualified Volume"] = df.apply(
        lambda row: calculate_qualified_volume(row["Shop Volume"], row["Site Volume"]),
        axis=1,
    )

    df["Qualified Points"] = df["Qualified Volume"].apply(calculate_qualified_points)
    df["Qualified Slab"] = df["Qualified Points"].apply(assign_slab)

    df["Lifting Frequency"] = (
        df[month_cols].gt(0).sum(axis=1) if month_cols else 0
    )
    df["Next Upgrade Slab"] = df["Qualified Slab"].apply(get_next_slab)
    df["Points to Next Slab"] = df.apply(
        lambda row: points_to_next(row["Qualified Points"], row["Qualified Slab"]),
        axis=1,
    )

    log_info(f"Data loading complete. Shape: {df.shape}")
    return df, month_cols


def _find_column(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    """Find the first matching column name from a list of candidates.

    Performs case-insensitive matching against DataFrame columns.

    Args:
        df: DataFrame to search.
        candidates: List of candidate column names in priority order.

    Returns:
        Matching column name from df.columns, or None.
    """
    col_lower_map = {c.lower().strip(): c for c in df.columns if isinstance(c, str)}
    for name in candidates:
        if name.lower().strip() in col_lower_map:
            return col_lower_map[name.lower().strip()]
    return None


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

    Creates a row of selectbox filters where each selection narrows the
    options available in subsequent dropdowns.

    Args:
        df: Input DataFrame to filter.
        key: Unique key prefix for widget state (use different keys per tab).
        filter_fields: Optional explicit list of column names to filter on.
            If None, uses default geo hierarchy + slab.

    Returns:
        Filtered DataFrame based on user selections.
    """
    if filter_fields is None:
        filter_fields = [
            col for col in ["Zone", "State", "Region", "District", "Distributor Name"]
            if col in df.columns
        ]
        if "Qualified Slab" in df.columns:
            filter_fields = ["Qualified Slab"] + filter_fields

    # Only keep fields that exist in df
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

def render_summary(df: pd.DataFrame, month_cols: list[str]) -> None:
    """Render the Summary tab with KPIs, slab cards, breakdown table and chart.

    Args:
        df: Full DataFrame.
        month_cols: List of month column names.
    """
    filtered = render_cascading_filters(df, key="summary")

    if filtered.empty:
        st.info("No data matches the selected filters.")
        return

    # --- KPI Row ---
    total_dealers = len(filtered)
    total_volume = filtered["Total Volume"].sum() if "Total Volume" in filtered.columns else 0
    total_qual_vol = filtered["Qualified Volume"].sum() if "Qualified Volume" in filtered.columns else 0
    total_points = filtered["Qualified Points"].sum() if "Qualified Points" in filtered.columns else 0

    kpi_cols = st.columns(4)
    kpis = [
        ("Total Dealers", format_indian(total_dealers)),
        ("Total Volume (MT)", format_indian(total_volume, decimal=1)),
        ("Qualified Volume (MT)", format_indian(total_qual_vol, decimal=1)),
        ("Total Qualified Points", format_indian(total_points)),
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
    if "Qualified Slab" in filtered.columns:
        slab_counts = filtered["Qualified Slab"].value_counts()
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

    # --- Slab Breakdown Table ---
    st.subheader("Slab Breakdown")
    summary_rows = []
    grand_count = 0
    grand_vol = 0.0
    grand_qual = 0.0
    grand_pts = 0.0

    for cfg in SLAB_CONFIG:
        slab_df = filtered[filtered["Qualified Slab"] == cfg["slab"]]
        count = len(slab_df)
        vol = slab_df["Total Volume"].sum() if "Total Volume" in slab_df.columns else 0
        qual = slab_df["Qualified Volume"].sum() if "Qualified Volume" in slab_df.columns else 0
        pts = slab_df["Qualified Points"].sum() if "Qualified Points" in slab_df.columns else 0
        grand_count += count
        grand_vol += vol
        grand_qual += qual
        grand_pts += pts
        summary_rows.append({
            "Slab": cfg["slab"],
            "Points Range": cfg["range"],
            "Dealer Count": count,
            "Total Volume": format_indian(vol, decimal=1),
            "Qualified Volume": format_indian(qual, decimal=1),
            "Total Points": format_indian(pts),
            "Gift": cfg["gift_full"],
        })

    summary_rows.append({
        "Slab": "TOTAL",
        "Points Range": "",
        "Dealer Count": grand_count,
        "Total Volume": format_indian(grand_vol, decimal=1),
        "Qualified Volume": format_indian(grand_qual, decimal=1),
        "Total Points": format_indian(grand_pts),
        "Gift": "",
    })

    summary_df = pd.DataFrame(summary_rows)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # --- Donut Chart: Slab Distribution ---
    if "Qualified Slab" in filtered.columns:
        st.subheader("Slab Distribution")
        slab_counts = filtered["Qualified Slab"].value_counts()
        labels = [s for s in SLAB_ORDER if s in slab_counts.index]
        values = [slab_counts[s] for s in labels]
        colors = [SLAB_COLORS[s] for s in labels]

        fig = go.Figure(
            data=[
                go.Pie(
                    labels=labels,
                    values=values,
                    marker=dict(colors=colors),
                    hole=0.4,
                    hovertemplate="<b>%{label}</b><br>Count: %{value}<br>%{percent}<extra></extra>",
                    textinfo="percent+label",
                )
            ]
        )
        fig.update_layout(
            plot_bgcolor="#ffffff",
            paper_bgcolor="#ffffff",
            margin=dict(l=20, r=20, t=20, b=20),
            height=400,
            showlegend=True,
        )
        st.plotly_chart(fig, use_container_width=True)


# ============================================================================
# SECTION 7 — TAB: DEALER DETAILS
# ============================================================================

def render_dealer_details(df: pd.DataFrame, month_cols: list[str]) -> None:
    """Render the Dealer Details tab with per-dealer table and filters.

    Filters: Distributor Name, State, Dealer Name, Qualified Slab.
    Columns: Dealer Name, Distributor Name, State, Region, Shop Volume,
             Site Volume, Total Volume, Qualified Volume, Qualified Slab,
             Qualified Points.

    Args:
        df: Full DataFrame.
        month_cols: List of month column names.
    """
    detail_filters = ["Distributor Name", "State", "Dealer Name", "Qualified Slab"]
    filtered = render_cascading_filters(df, key="dealer_detail", filter_fields=detail_filters)

    if filtered.empty:
        st.info("No data matches the selected filters.")
        return

    # --- Data Table ---
    display_cols = [
        c for c in [
            "Dealer Name", "Distributor Name", "State", "Region",
            "Shop Volume", "Site Volume", "Total Volume",
            "Qualified Volume", "Qualified Slab", "Qualified Points",
        ]
        if c in filtered.columns
    ]

    st.subheader(f"Dealer Details ({len(filtered)} records)")

    def _highlight_total_row(row: pd.Series) -> list[str]:
        """Apply bold grey background to total/summary rows."""
        for field in ["Dealer Name", "Distributor Name"]:
            val = row.get(field)
            if isinstance(val, str) and "total" in val.lower():
                return ["font-weight: 700; background-color: #f1f5f9"] * len(row)
        return [""] * len(row)

    styled = filtered[display_cols].style.apply(_highlight_total_row, axis=1)
    st.dataframe(styled, use_container_width=True, hide_index=True, height=500)


# ============================================================================
# SECTION 8 — MAIN
# ============================================================================

def main() -> None:
    """Entry point for the Streamlit dashboard application."""
    st.set_page_config(
        page_title="Scheme Dashboard",
        layout="wide",
        page_icon="📊",
    )

    inject_custom_css()

    # --- Header Bar ---
    st.markdown(
        """
        <div class="header-bar">
            <h1>📊 Scheme Dashboard</h1>
            <span>Slab Analysis & Dealer Tracker</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --- Load Data ---
    df, month_cols = load_data()

    # --- File info ---
    file_path = _find_excel_file()
    st.caption(f"Data source: `{file_path.name}` — {len(df)} records loaded")

    # --- Tabs ---
    tab_summary, tab_details = st.tabs([
        "📊 Summary",
        "🔍 Dealer Details",
    ])

    with tab_summary:
        render_summary(df, month_cols)

    with tab_details:
        render_dealer_details(df, month_cols)


if __name__ == "__main__":
    main()
