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
- `app.py` — Single-file Streamlit app (V1), card-based white UI with shadows
- `app_v2.py` — Minimalist clean redesign (V2), Swiss-style with Inter font
- `src/lib/logger.py` — Simple print-based logger (info/error)
- `scripts/` — Standalone report generators (no Streamlit dependency)
- `blueprints/` — Task SOPs in markdown
- `tests/unit/` and `tests/integration/` — Test suites
- `data/` — Source Excel files (.xlsx)
- `.workspace/` — Temp output files (gitignored)
- `.claude/skills/ui-ux-pro-max/` — UI/UX design intelligence skill

## Key Patterns
- **SLAB_CONFIG** is single source of truth for all tier definitions
- All slab-derived dicts are computed from SLAB_CONFIG, never hardcoded separately
- Cascading filters reuse `render_cascading_filters()` with unique key prefixes per tab
- Indian number formatting (`format_indian`) for all displayed values
- `@st.cache_data` on all data-loading functions
- Both app.py and app_v2.py share the same data pipeline (SLAB_CONFIG, load_data, helpers)
- New dashboard versions are separate files, not modifications of existing ones
- Use ui-ux-pro-max skill (`search.py --design-system`) for design decisions

## Error Protocol
1. Log the error via `src/lib/logger.py`
2. Show user-friendly message via `st.error()` or `st.warning()`
3. Never let the app crash silently — always `st.stop()` on fatal errors
4. Record resolution in LEARNINGS.md

## CSS Design Systems

### V1 (app.py) — Card-based White UI
- Card-based white UI with subtle shadows and dark gradient header
- Color-coded elements per slab via SLAB_COLORS dict (5px left borders)
- Typography: .kpi-label (small, uppercase, muted), .kpi-value (bold, dark)
- Zebra-striped tables with dark (#1e293b) header backgrounds
- Responsive layout via st.columns()

### V2 (app_v2.py) — Minimalist Clean
- Swiss Minimalism: no shadows, no gradients, generous whitespace
- Inter font via Google Fonts (@import in inject_minimalist_css)
- Colors: #111111 text, #6B7280 muted, #2563EB accent, #E5E7EB borders
- KPI metrics: borderless with blue accent divider (v2-kpi, v2-kpi-value)
- Slab indicators: 8px colored dot (v2-slab-dot) instead of thick borders
- Tables: #F9FAFB headers, no zebra striping, subtle hover only
- Widget keys prefixed with `v2_` to avoid state conflicts with V1
- CSS classes prefixed with `v2-` (v2-header, v2-kpi, v2-slab, v2-section-title)

## Deployment
- Deployed on Streamlit Community Cloud from `base` branch
- Main file: `app_v2.py`
- Streamlit Cloud may not show all branches in dropdown — use well-known branches like `base`
- Merge feature branches into `base` via GitHub PR before deploying
