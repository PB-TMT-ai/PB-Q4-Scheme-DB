# CLAUDE.md — Project Instructions

## Project Overview
Streamlit-based dashboard for scheme/slab analysis with Excel data ingestion,
cascading filters, and Plotly visualizations.

## Coding Standards
- **Python 3.10+** with type hints on all function signatures
- **Google-style docstrings** on all public functions
- `snake_case` for functions/variables, `PascalCase` for classes
- `f-strings` only — no `.format()` or `%` formatting
- `pathlib.Path` over `os.path` everywhere
- No bare `except` — catch specific exceptions (ValueError, FileNotFoundError, etc.)
- `dataclasses` or Pydantic for data models if needed

## Project Structure
- `app.py` — Single-file Streamlit app, sectioned with numbered comments
- `src/lib/logger.py` — Simple print-based logger (info/error)
- `scripts/` — Standalone report generators (no Streamlit dependency)
- `blueprints/` — Task SOPs in markdown
- `tests/unit/` and `tests/integration/` — Test suites
- `data/` — Source Excel files (.xlsx)
- `.workspace/` — Temp output files (gitignored)

## Key Patterns
- **SLAB_CONFIG** is single source of truth for all tier definitions
- All slab-derived dicts are computed from SLAB_CONFIG, never hardcoded separately
- Cascading filters reuse `render_cascading_filters()` with unique key prefixes per tab
- Extra filters (e.g. Vol. to Achieve) should be rendered **inline** in the same `st.columns()` row as cascading filters, not in a separate row
- Indian number formatting (`format_indian`) for all displayed values
- `@st.cache_data` on all data-loading functions

## Deployment
- Hosted on **Streamlit Cloud**
- Deployment branch: `claude/dealer-slab-color-coding-dFor3` (verify in Streamlit Cloud settings)
- Always confirm the deployment branch before pushing fixes
- The GitHub default branch (`claude/streamlit-dashboard-project-1WQZy`) differs from the Streamlit Cloud deployment branch — **do not assume they are the same**
- After merging a PR, verify the Streamlit app has redeployed (hard refresh with Ctrl+Shift+R)

## Branch Management — Critical
- Features are spread across multiple branches with **different states** (e.g. 2-tab vs 4-tab versions)
- Before making any change, **always check the deployment branch first** (`git fetch origin claude/dealer-slab-color-coding-dFor3`) and branch off from it
- Protected branches cannot be pushed to directly — changes go through PRs
- Only branches matching `claude/*-<sessionId>` can be pushed from Claude Code sessions

## Gotchas & Pitfalls
- **`render_cascading_filters()` is only for standard column-value filters.** Computed/range-based filters (like Vol. to Achieve) need manual `st.selectbox` with custom filtering logic — do NOT try to add them to `filter_fields`
- **Vol. to Achieve is computed, not a data column.** It's calculated at display time via `_calc_vol_to_achieve()` using `NEXT_SLAB_VOLUME` dict. The filter uses range buckets (0-20 MT, 21-40 MT, etc.), not exact values
- **Streamlit widget keys must be globally unique.** When rendering filters manually (not via `render_cascading_filters`), use `key=f"{tab_prefix}_{field}"` pattern to avoid `DuplicateWidgetID` errors
- **`display_df` column names differ from `filtered` DataFrame.** Columns are renamed for display (e.g. `Shop Volume` → `Qual. Shop Vol.`). When computing on `display_df`, use the **renamed** column names
- **All numbers displayed as whole integers** — no decimals anywhere (volumes, points, percentages). Use `.round(0).astype(int)` for table columns, `decimal=0` in `format_indian()`
- **Self-counter dealers are excluded during data loading** (filtered out in `load_data()`). The dealer count shown is post-exclusion
- **SLAB_CONFIG changes must be mirrored** in `scripts/generate_slab_report.py` and validated in `tests/unit/test_helpers.py` (required-keys check)

## Error Protocol
1. Log the error via `src/lib/logger.py`
2. Show user-friendly message via `st.error()` or `st.warning()`
3. Never let the app crash silently — always `st.stop()` on fatal errors
4. Record resolution in LEARNINGS.md

## CSS Design System
- Card-based white UI with subtle shadows
- Color-coded elements per slab via SLAB_COLORS dict
- Typography: .kpi-label (small, uppercase, muted), .kpi-value (bold, dark)
- Responsive layout via st.columns()
