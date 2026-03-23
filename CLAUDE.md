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
- `app_v2.py` — Bento Grid redesign (V2), Apple-style with Fira Sans/Code fonts
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

### V2 (app_v2.py) — Bento Grid
- Apple-style modular cards: mixed sizes (2x1 hero, 1x1 standard), rounded corners (20px)
- Fira Sans (headings) + Fira Code (numbers) via Google Fonts (@import in inject_bento_css)
- Page background: #F5F5F7 (Apple off-white), card background: #FFFFFF with soft shadow
- KPI metrics: bento cards with left-aligned Fira Code numbers (v2-bento-card, v2-kpi)
- Slab tiles: colored accent bar at top (v2-slab-accent), hover scale 1.02
- Tables: rounded 12px, #F5F5F7 headers, no borders, subtle hover
- Tabs: pill-style with rounded background instead of underline
- Widget keys prefixed with `v2_` to avoid state conflicts with V1
- CSS classes prefixed with `v2-` (v2-header, v2-bento-card, v2-slab, v2-section-title)

## Deployment
- Deployed on Streamlit Community Cloud from `base` branch
- Main file: `app_v2.py`
- Streamlit Cloud may not show all branches in dropdown — use well-known branches like `base`
- Merge feature branches into `base` via GitHub PR before deploying
