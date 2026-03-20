"""
Streamlit Dashboard — Scheme Slab Analysis
==========================================
Single-file Streamlit app structured in numbered sections.
Place your Excel data file in the data/ directory and configure
SLAB_CONFIG to match your scheme tiers.

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
# Customize this list for your scheme. Each dict defines one slab/tier.
# Fields:
#   slab        — Short label displayed in UI
#   range       — Human-readable volume range string
#   lower       — Inclusive lower bound (volume)
#   upper       — Exclusive upper bound (use float("inf") for the top tier)
#   gift        — Short gift/reward label
#   gift_full   — Full gift description
#   category    — Grouping category (e.g. "Bronze", "Silver", "Gold")
#   value       — Monetary value of the gift/reward
#   value_tds   — Value after TDS deduction
#   color       — Hex color for UI elements

SLAB_CONFIG: list[dict] = [
    {
        "slab": "Slab 1",
        "range": "0 – 499",
        "lower": 0,
        "upper": 500,
        "gift": "No Gift",
        "gift_full": "Below minimum qualification",
        "category": "Unqualified",
        "value": 0,
        "value_tds": 0,
        "color": "#94a3b8",
    },
    {
        "slab": "Slab 2",
        "range": "500 – 999",
        "lower": 500,
        "upper": 1000,
        "gift": "Bronze Gift",
        "gift_full": "Bronze tier reward",
        "category": "Bronze",
        "value": 5000,
        "value_tds": 4500,
        "color": "#f59e0b",
    },
    {
        "slab": "Slab 3",
        "range": "1,000 – 2,499",
        "lower": 1000,
        "upper": 2500,
        "gift": "Silver Gift",
        "gift_full": "Silver tier reward",
        "category": "Silver",
        "value": 15000,
        "value_tds": 13500,
        "color": "#6366f1",
    },
    {
        "slab": "Slab 4",
        "range": "2,500 – 4,999",
        "lower": 2500,
        "upper": 5000,
        "gift": "Gold Gift",
        "gift_full": "Gold tier reward",
        "category": "Gold",
        "value": 35000,
        "value_tds": 31500,
        "color": "#10b981",
    },
    {
        "slab": "Slab 5",
        "range": "5,000+",
        "lower": 5000,
        "upper": float("inf"),
        "gift": "Platinum Gift",
        "gift_full": "Platinum tier reward",
        "category": "Platinum",
        "value": 75000,
        "value_tds": 67500,
        "color": "#3b82f6",
    },
]

# Derived lookup dictionaries — all computed from SLAB_CONFIG
SLAB_GIFT_MAP: dict[str, str] = {s["slab"]: s["gift_full"] for s in SLAB_CONFIG}
SLAB_VALUE_TDS: dict[str, int] = {s["slab"]: s["value_tds"] for s in SLAB_CONFIG}
SLAB_COLORS: dict[str, str] = {s["slab"]: s["color"] for s in SLAB_CONFIG}
SLAB_ORDER: list[str] = [s["slab"] for s in SLAB_CONFIG]

# Next-slab mapping (each slab → next higher slab, top tier → None)
NEXT_SLAB_MAP: dict[str, Optional[str]] = {}
for _i, _cfg in enumerate(SLAB_CONFIG):
    NEXT_SLAB_MAP[_cfg["slab"]] = (
        SLAB_CONFIG[_i + 1]["slab"] if _i + 1 < len(SLAB_CONFIG) else None
    )

# Threshold to reach next slab (lower bound of next slab)
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


def assign_slab(vol: float) -> str:
    """Assign a slab label based on volume using SLAB_CONFIG thresholds.

    Args:
        vol: The volume value to classify.

    Returns:
        Slab label string (e.g. "Slab 3").
    """
    if pd.isna(vol):
        vol = 0.0
    vol = float(vol)
    for cfg in SLAB_CONFIG:
        if cfg["lower"] <= vol < cfg["upper"]:
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


def volume_to_next(vol: float, current: str) -> Optional[float]:
    """Calculate the volume gap to reach the next slab tier.

    Args:
        vol: Current total volume.
        current: Current slab label.

    Returns:
        Volume needed to reach next tier, or None if at max slab.
    """
    threshold = NEXT_SLAB_THRESHOLD.get(current)
    if threshold is None:
        return None
    gap = threshold - float(vol)
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
    labels, coerces numerics, cleans text fields, and derives slab columns.

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
    # Attempt to map datetime or string month headers to short labels
    col_map: dict[str, str] = {}
    month_cols: list[str] = []

    for col in df.columns:
        matched = False
        # Try parsing as datetime
        if hasattr(col, "strftime"):
            short = col.strftime("%b")
            if short in MONTH_LABELS and short not in col_map.values():
                col_map[col] = short
                month_cols.append(short)
                matched = True
        # Try matching string names
        if not matched and isinstance(col, str):
            col_str = col.strip()
            for label in MONTH_LABELS:
                if col_str.lower().startswith(label.lower()) and label not in col_map.values():
                    col_map[col] = label
                    month_cols.append(label)
                    break

    if col_map:
        df = df.rename(columns=col_map)

    # If no month columns detected, look for columns already named as months
    if not month_cols:
        for label in MONTH_LABELS:
            if label in df.columns:
                month_cols.append(label)

    log_info(f"Month columns identified: {month_cols}")

    # --- Coerce month columns to numeric ---
    for col in month_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # --- Coerce common aggregate columns to numeric ---
    for col in df.columns:
        if isinstance(col, str) and any(
            kw in col.lower()
            for kw in ["total", "volume", "qty", "quantity", "amount", "value"]
        ):
            if col not in month_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # --- Clean text columns ---
    text_cols = ["State", "Zone", "District", "Distributor Name", "Region"]
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

    # --- Derived columns ---
    # Calculate total volume across month columns if not already present
    if month_cols:
        if "Total Volume" not in df.columns:
            df["Total Volume"] = df[month_cols].sum(axis=1)
        vol_col = "Total Volume"
    else:
        # Fallback: look for any column with "total" in name
        total_candidates = [
            c for c in df.columns
            if isinstance(c, str) and "total" in c.lower()
        ]
        vol_col = total_candidates[0] if total_candidates else None

    if vol_col and vol_col in df.columns:
        df["Qualified Slab"] = df[vol_col].apply(assign_slab)
        df["Lifting Frequency"] = (
            df[month_cols].gt(0).sum(axis=1) if month_cols else 0
        )
        df["Next Upgrade Slab"] = df["Qualified Slab"].apply(get_next_slab)
        df["Vol to Next Slab"] = df.apply(
            lambda row: volume_to_next(row[vol_col], row["Qualified Slab"]),
            axis=1,
        )
    else:
        log_info("No volume column found — skipping derived slab columns")

    log_info(f"Data loading complete. Shape: {df.shape}")
    return df, month_cols


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


def render_cascading_filters(df: pd.DataFrame, key: str) -> pd.DataFrame:
    """Render cascading dropdown filters and return the filtered DataFrame.

    Creates a row of selectbox filters where each selection narrows the
    options available in subsequent dropdowns.

    Args:
        df: Input DataFrame to filter.
        key: Unique key prefix for widget state (use different keys per tab).

    Returns:
        Filtered DataFrame based on user selections.
    """
    filter_fields = [
        col for col in ["Zone", "State", "Region", "District", "Distributor Name"]
        if col in df.columns
    ]

    if not filter_fields:
        return df

    # Add slab filter if available
    if "Qualified Slab" in df.columns:
        filter_fields = ["Qualified Slab"] + filter_fields

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
# SECTION 6 — TAB: OVERVIEW
# ============================================================================

def render_overview(df: pd.DataFrame, month_cols: list[str]) -> None:
    """Render the Overview tab with KPI cards and slab distribution.

    Args:
        df: Full DataFrame.
        month_cols: List of month column names.
    """
    filtered = render_cascading_filters(df, key="overview")

    if filtered.empty:
        st.info("No data matches the selected filters.")
        return

    # --- KPI Row ---
    total_distributors = len(filtered)
    total_volume = filtered["Total Volume"].sum() if "Total Volume" in filtered.columns else 0
    avg_volume = filtered["Total Volume"].mean() if "Total Volume" in filtered.columns else 0
    avg_frequency = filtered["Lifting Frequency"].mean() if "Lifting Frequency" in filtered.columns else 0

    kpi_cols = st.columns(4)
    kpis = [
        ("Total Distributors", format_indian(total_distributors)),
        ("Total Volume", format_indian(total_volume)),
        ("Avg Volume", format_indian(avg_volume, decimal=1)),
        ("Avg Lifting Months", f"{avg_frequency:.1f}"),
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
                        <div class="slab-label">{cfg['slab']} — {cfg['range']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    st.markdown("<br>", unsafe_allow_html=True)

    # --- Monthly Volume Trend Chart ---
    if month_cols:
        st.subheader("Monthly Volume Trend")
        monthly_totals = filtered[month_cols].sum()

        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                x=month_cols,
                y=monthly_totals.values,
                marker_color="#3b82f6",
                hovertemplate="<b>%{x}</b><br>Volume: %{y:,.0f}<extra></extra>",
            )
        )
        fig.update_layout(
            plot_bgcolor="#ffffff",
            paper_bgcolor="#ffffff",
            margin=dict(l=40, r=20, t=20, b=40),
            xaxis=dict(title="Month", showgrid=False),
            yaxis=dict(title="Volume", gridcolor="#f1f5f9"),
            height=350,
        )
        st.plotly_chart(fig, use_container_width=True)


# ============================================================================
# SECTION 7 — TAB: SLAB ANALYSIS
# ============================================================================

def render_slab_analysis(df: pd.DataFrame, month_cols: list[str]) -> None:
    """Render the Slab Analysis tab with distribution chart and breakdown.

    Args:
        df: Full DataFrame.
        month_cols: List of month column names.
    """
    filtered = render_cascading_filters(df, key="slab_analysis")

    if filtered.empty:
        st.info("No data matches the selected filters.")
        return

    if "Qualified Slab" not in filtered.columns:
        st.warning("Slab assignment not available. Check data columns.")
        return

    col1, col2 = st.columns([1, 1])

    # --- Pie Chart ---
    with col1:
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
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True)

    # --- Slab Summary Table ---
    with col2:
        st.subheader("Slab Breakdown")
        summary_rows = []
        for cfg in SLAB_CONFIG:
            slab_df = filtered[filtered["Qualified Slab"] == cfg["slab"]]
            count = len(slab_df)
            vol = slab_df["Total Volume"].sum() if "Total Volume" in slab_df.columns else 0
            summary_rows.append({
                "Slab": cfg["slab"],
                "Range": cfg["range"],
                "Count": count,
                "Total Volume": format_indian(vol),
                "Gift": cfg["gift"],
                "Value (post-TDS)": format_indian(cfg["value_tds"], prefix="₹"),
            })

        summary_df = pd.DataFrame(summary_rows)
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

    # --- Upgrade Potential ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Upgrade Potential")
    upgrade_df = filtered[filtered["Vol to Next Slab"].notna() & (filtered["Vol to Next Slab"] > 0)].copy()
    if not upgrade_df.empty and "Vol to Next Slab" in upgrade_df.columns:
        upgrade_df = upgrade_df.sort_values("Vol to Next Slab", ascending=True).head(20)
        display_cols = [
            c for c in ["Distributor Name", "State", "District", "Total Volume",
                        "Qualified Slab", "Next Upgrade Slab", "Vol to Next Slab"]
            if c in upgrade_df.columns
        ]
        st.dataframe(
            upgrade_df[display_cols].reset_index(drop=True),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("No distributors with upgrade potential in the current filter.")


# ============================================================================
# SECTION 8 — TAB: DISTRIBUTOR DETAIL
# ============================================================================

def render_distributor_detail(df: pd.DataFrame, month_cols: list[str]) -> None:
    """Render the Distributor Detail tab with per-distributor drill-down.

    Args:
        df: Full DataFrame.
        month_cols: List of month column names.
    """
    filtered = render_cascading_filters(df, key="dist_detail")

    if filtered.empty:
        st.info("No data matches the selected filters.")
        return

    # Display the filtered table with styling
    display_cols = [
        c for c in [
            "Distributor Name", "State", "District", "Zone",
            "Total Volume", "Qualified Slab", "Lifting Frequency",
            "Next Upgrade Slab", "Vol to Next Slab",
        ] + month_cols
        if c in filtered.columns
    ]

    st.subheader(f"Distributors ({len(filtered)})")

    def _highlight_total_row(row: pd.Series) -> list[str]:
        """Apply bold grey background to total/summary rows."""
        if isinstance(row.get("Distributor Name"), str) and "total" in row["Distributor Name"].lower():
            return ["font-weight: 700; background-color: #f1f5f9"] * len(row)
        return [""] * len(row)

    styled = filtered[display_cols].style.apply(_highlight_total_row, axis=1)
    st.dataframe(styled, use_container_width=True, hide_index=True, height=500)

    # --- Individual distributor volume chart ---
    if "Distributor Name" in filtered.columns and month_cols:
        st.markdown("<br>", unsafe_allow_html=True)
        dist_names = sorted(filtered["Distributor Name"].unique())
        if dist_names:
            selected_dist = st.selectbox(
                "Select distributor for monthly trend:",
                dist_names,
                key="dist_detail_select",
            )
            dist_row = filtered[filtered["Distributor Name"] == selected_dist]
            if not dist_row.empty and month_cols:
                vals = dist_row[month_cols].iloc[0].values

                fig = go.Figure()
                fig.add_trace(
                    go.Scatter(
                        x=month_cols,
                        y=vals,
                        mode="lines+markers",
                        line=dict(color="#6366f1", width=2),
                        marker=dict(size=8),
                        hovertemplate="<b>%{x}</b><br>Volume: %{y:,.0f}<extra></extra>",
                    )
                )
                fig.update_layout(
                    plot_bgcolor="#ffffff",
                    paper_bgcolor="#ffffff",
                    margin=dict(l=40, r=20, t=30, b=40),
                    xaxis=dict(title="Month", showgrid=False),
                    yaxis=dict(title="Volume", gridcolor="#f1f5f9"),
                    height=300,
                    title=dict(text=f"Monthly Trend — {selected_dist}", font=dict(size=14)),
                )
                st.plotly_chart(fig, use_container_width=True)


# ============================================================================
# SECTION 9 — TAB: GIFT SUMMARY
# ============================================================================

def render_gift_summary(df: pd.DataFrame, month_cols: list[str]) -> None:
    """Render the Gift Summary tab with reward value calculations.

    Args:
        df: Full DataFrame.
        month_cols: List of month column names.
    """
    filtered = render_cascading_filters(df, key="gift_summary")

    if filtered.empty:
        st.info("No data matches the selected filters.")
        return

    if "Qualified Slab" not in filtered.columns:
        st.warning("Slab data not available.")
        return

    # --- Gift value summary ---
    st.subheader("Gift Value Summary by Slab")

    summary_rows = []
    total_count = 0
    total_value = 0
    total_tds = 0

    for cfg in SLAB_CONFIG:
        slab_df = filtered[filtered["Qualified Slab"] == cfg["slab"]]
        count = len(slab_df)
        value = count * cfg["value"]
        tds = count * cfg["value_tds"]
        total_count += count
        total_value += value
        total_tds += tds
        summary_rows.append({
            "Slab": cfg["slab"],
            "Gift": cfg["gift_full"],
            "Count": count,
            "Unit Value": format_indian(cfg["value"], prefix="₹"),
            "Total Value": format_indian(value, prefix="₹"),
            "Total (post-TDS)": format_indian(tds, prefix="₹"),
        })

    # Add total row
    summary_rows.append({
        "Slab": "TOTAL",
        "Gift": "",
        "Count": total_count,
        "Unit Value": "",
        "Total Value": format_indian(total_value, prefix="₹"),
        "Total (post-TDS)": format_indian(total_tds, prefix="₹"),
    })

    summary_df = pd.DataFrame(summary_rows)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    # --- Stacked bar chart of gift values ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Gift Value Distribution")

    chart_slabs = [cfg["slab"] for cfg in SLAB_CONFIG if cfg["value"] > 0]
    chart_counts = [
        len(filtered[filtered["Qualified Slab"] == s])
        for s in chart_slabs
    ]
    chart_values = [
        len(filtered[filtered["Qualified Slab"] == s]) * next(
            c["value_tds"] for c in SLAB_CONFIG if c["slab"] == s
        )
        for s in chart_slabs
    ]
    chart_colors = [SLAB_COLORS[s] for s in chart_slabs]

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=chart_slabs,
            y=chart_values,
            marker_color=chart_colors,
            text=[format_indian(v, prefix="₹") for v in chart_values],
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Value: %{text}<br>Count: %{customdata}<extra></extra>",
            customdata=chart_counts,
        )
    )
    fig.update_layout(
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        margin=dict(l=40, r=20, t=20, b=40),
        xaxis=dict(title="Slab", showgrid=False),
        yaxis=dict(title="Total Gift Value (post-TDS)", gridcolor="#f1f5f9"),
        height=350,
    )
    st.plotly_chart(fig, use_container_width=True)


# ============================================================================
# SECTION 10 — MAIN
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
            <span>Slab Analysis & Gift Tracker</span>
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
    tab_overview, tab_slab, tab_detail, tab_gift = st.tabs([
        "📈 Overview",
        "🏷️ Slab Analysis",
        "🔍 Distributor Detail",
        "🎁 Gift Summary",
    ])

    with tab_overview:
        render_overview(df, month_cols)

    with tab_slab:
        render_slab_analysis(df, month_cols)

    with tab_detail:
        render_distributor_detail(df, month_cols)

    with tab_gift:
        render_gift_summary(df, month_cols)


if __name__ == "__main__":
    main()
