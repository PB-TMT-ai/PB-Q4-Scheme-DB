# Blueprint: Add a New Dashboard Tab

## Overview
Standard operating procedure for adding a new tab to the Streamlit dashboard.

## Steps

### 1. Define the Tab Renderer Function
Create a new function in app.py following the pattern:

```python
# ============================================================================
# SECTION N — TAB: <TAB_NAME>
# ============================================================================

def render_<tab_name>(df: pd.DataFrame) -> None:
    """Render the <Tab Name> tab.

    Args:
        df: Full DataFrame.
    """
    filtered = render_cascading_filters(df, key="<tab_key>")

    if filtered.empty:
        st.info("No data matches the selected filters.")
        return

    # Tab content here...
```

### 2. Register the Tab in main()
Add the tab to the `st.tabs()` call and wire up the renderer:

```python
tabs = st.tabs([..., "🆕 New Tab"])
# ...
with tab_new:
    render_<tab_name>(df)
```

### 3. Key Rules
- Use a unique `key` prefix for `render_cascading_filters` (e.g. `"new_tab"`)
- Always handle empty state with `st.info()` + early return
- Use `format_indian()` for all displayed numbers
- Use Plotly `graph_objects` for charts (not `plotly.express`)
- Follow the CSS design system (`.kpi-card`, `.slab-card` classes)

### 4. Available Data Columns
The DataFrame includes these standard columns (mapped from Excel):
- `Dealer Name`, `Distributor Name`, `State`, `District`, `Zone`
- `Shop Volume`, `Site Volume`, `Total Volume`, `Qualified Volume`
- `Qualified Points`, `Qualified Slab`, `Q4 Volume`
- `Next Upgrade Slab`, `Points to Next Slab`

### 5. Test
- Verify filters cascade correctly
- Verify empty state displays info message
- Verify charts render with data
