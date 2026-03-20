"""Unit tests for helper functions defined in app.py."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from app import (
    NEXT_SLAB_MAP,
    SLAB_CONFIG,
    SLAB_ORDER,
    assign_slab,
    format_indian,
    get_next_slab,
    volume_to_next,
)


# ---------------------------------------------------------------------------
# format_indian
# ---------------------------------------------------------------------------

class TestFormatIndian:
    """Tests for the format_indian() function."""

    def test_zero(self) -> None:
        assert format_indian(0) == "0"

    def test_hundreds(self) -> None:
        assert format_indian(500) == "500"

    def test_thousands(self) -> None:
        assert format_indian(1234) == "1,234"

    def test_lakhs(self) -> None:
        assert format_indian(123456) == "1,23,456"

    def test_crores(self) -> None:
        assert format_indian(12345678) == "1,23,45,678"

    def test_prefix(self) -> None:
        assert format_indian(5000, prefix="₹") == "₹5,000"

    def test_decimal(self) -> None:
        result = format_indian(1234.567, decimal=2)
        assert result == "1,234.57"

    def test_negative(self) -> None:
        result = format_indian(-5000)
        assert result == "-5,000"

    def test_nan(self) -> None:
        result = format_indian(float("nan"))
        assert result == "0"

    def test_large_number(self) -> None:
        result = format_indian(1234567890)
        assert result == "1,23,45,67,890"


# ---------------------------------------------------------------------------
# assign_slab
# ---------------------------------------------------------------------------

class TestAssignSlab:
    """Tests for the assign_slab() function."""

    def test_zero_volume(self) -> None:
        assert assign_slab(0) == "Slab 1"

    def test_boundary_lower(self) -> None:
        assert assign_slab(500) == "Slab 2"

    def test_boundary_just_below(self) -> None:
        assert assign_slab(499) == "Slab 1"

    def test_mid_range(self) -> None:
        assert assign_slab(1500) == "Slab 3"

    def test_top_slab(self) -> None:
        assert assign_slab(10000) == "Slab 5"

    def test_nan_input(self) -> None:
        assert assign_slab(float("nan")) == "Slab 1"


# ---------------------------------------------------------------------------
# get_next_slab
# ---------------------------------------------------------------------------

class TestGetNextSlab:
    """Tests for the get_next_slab() function."""

    def test_first_slab(self) -> None:
        assert get_next_slab("Slab 1") == "Slab 2"

    def test_last_slab(self) -> None:
        assert get_next_slab("Slab 5") is None

    def test_unknown_slab(self) -> None:
        assert get_next_slab("Unknown") is None


# ---------------------------------------------------------------------------
# volume_to_next
# ---------------------------------------------------------------------------

class TestVolumeToNext:
    """Tests for the volume_to_next() function."""

    def test_gap_exists(self) -> None:
        result = volume_to_next(300, "Slab 1")
        assert result == 200.0

    def test_already_past_threshold(self) -> None:
        result = volume_to_next(600, "Slab 1")
        assert result == 0.0

    def test_max_slab_returns_none(self) -> None:
        result = volume_to_next(10000, "Slab 5")
        assert result is None


# ---------------------------------------------------------------------------
# SLAB_CONFIG integrity
# ---------------------------------------------------------------------------

class TestSlabConfig:
    """Tests for SLAB_CONFIG consistency."""

    def test_slabs_sorted_by_lower_bound(self) -> None:
        lowers = [s["lower"] for s in SLAB_CONFIG]
        assert lowers == sorted(lowers)

    def test_no_gaps_in_bounds(self) -> None:
        for i in range(len(SLAB_CONFIG) - 1):
            assert SLAB_CONFIG[i]["upper"] == SLAB_CONFIG[i + 1]["lower"]

    def test_all_required_keys(self) -> None:
        required = {"slab", "range", "lower", "upper", "gift", "gift_full",
                     "category", "value", "value_tds", "color"}
        for cfg in SLAB_CONFIG:
            assert required.issubset(cfg.keys()), f"Missing keys in {cfg['slab']}"

    def test_slab_order_matches_config(self) -> None:
        assert SLAB_ORDER == [s["slab"] for s in SLAB_CONFIG]
