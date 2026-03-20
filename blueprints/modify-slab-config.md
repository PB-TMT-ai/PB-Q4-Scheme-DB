# Blueprint: Modify Slab Configuration

## Overview
How to add, remove, or change slab/tier definitions in the dashboard.

## Steps

### 1. Edit SLAB_CONFIG in app.py (Section 1)
Add or modify entries in the `SLAB_CONFIG` list. Each entry requires:

| Field       | Type        | Description                          |
|-------------|-------------|--------------------------------------|
| slab        | str         | Short label (e.g. "Slab 6")         |
| range       | str         | Human-readable range                 |
| lower       | float       | Inclusive lower bound                |
| upper       | float       | Exclusive upper bound (inf for top)  |
| gift        | str         | Short gift label                     |
| gift_full   | str         | Full gift description                |
| category    | str         | Grouping category                    |
| value       | int         | Gift monetary value                  |
| value_tds   | int         | Value after TDS                      |
| color       | str         | Hex color for UI elements            |

### 2. Rules
- **Bounds must be contiguous**: `upper` of slab N must equal `lower` of slab N+1
- **Last slab** must have `upper: float("inf")`
- **Colors** should be visually distinct
- All derived dicts (SLAB_GIFT_MAP, SLAB_COLORS, etc.) auto-update

### 3. Mirror in Scripts
Update `SLAB_CONFIG` in `scripts/generate_slab_report.py` to match.

### 4. Validate
Run unit tests to verify config integrity:
```bash
pytest tests/unit/test_helpers.py::TestSlabConfig -v
```
