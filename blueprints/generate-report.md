# Blueprint: Generate an Excel Report

## Overview
How to create standalone report generation scripts.

## Steps

### 1. Create Script in scripts/
```
scripts/generate_<report_name>.py
```

### 2. Script Structure
```python
"""
<Report description>
"""
from __future__ import annotations
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.lib.logger import info as log_info, error as log_error

# Mirror SLAB_CONFIG and helpers from app.py
# Load data independently (no Streamlit)
# Generate Excel with openpyxl styling
# Output to .workspace/

def main() -> None:
    ...

if __name__ == "__main__":
    main()
```

### 3. Excel Styling Standards
- **Headers**: Bold white text on blue fill (#1F4E79)
- **Auto-width columns**: Max 40 characters
- **Frozen panes**: Freeze row 1 (headers)
- **Tab colors**: Use hex colors for sheet tabs
- **All numbers**: Whole integers (`int(round(...))`) — no decimals
- **Output**: Always to `.workspace/` directory

### 4. Current Report Sheets (generate_slab_report.py)
| # | Sheet Name | Tab Color | Content |
|---|-----------|-----------|---------|
| 1 | Slab Summary | #1F4E79 | Per-slab dealer count, volumes, points, gift |
| 2 | Dealer Detail | #10B981 | Per-dealer breakdown with Vol. to Achieve |
| 3 | Near-Upgrade | #F59E0B | Dealers within 500 pts of next slab |
| 4 | Zone Performance | #10B981 | Zone-level aggregation + slab distribution |
| 5 | State Performance | #3B82F6 | State-level aggregation + slab distribution |
| 6 | Distributor Performance | #6366F1 | Distributor-level aggregation + slab distribution |

Sheets 4–6 use the `_write_perf_sheet()` helper to avoid duplication.

### 5. Run
```bash
python scripts/generate_<report_name>.py
```
