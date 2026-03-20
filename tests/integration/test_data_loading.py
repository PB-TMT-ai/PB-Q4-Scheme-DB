"""Integration tests for data loading.

These tests require a valid Excel file in data/ to run.
Skip gracefully if no data file is present.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

DATA_DIR = Path(__file__).resolve().parent.parent.parent / "data"


@pytest.fixture
def has_data_file() -> bool:
    """Check if a data file exists for integration tests."""
    files = list(DATA_DIR.glob("*.xlsx"))
    if not files:
        pytest.skip("No Excel file in data/ — skipping integration tests")
    return True


class TestDataLoading:
    """Integration tests for data loading pipeline."""

    def test_report_script_imports(self) -> None:
        """Verify the report script can be imported without errors."""
        from scripts.generate_slab_report import SLAB_CONFIG, assign_slab, format_indian
        assert len(SLAB_CONFIG) > 0
        assert assign_slab(0) == "Slab 1"

    def test_load_data_with_file(self, has_data_file: bool) -> None:
        """Test data loading with a real file."""
        from scripts.generate_slab_report import load_data
        df, month_cols = load_data()
        assert len(df) > 0
        assert isinstance(month_cols, list)
