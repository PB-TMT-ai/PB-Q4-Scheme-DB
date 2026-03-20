"""Unit tests for helper functions defined in app.py."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from app import (
    NEXT_SLAB_MAP,
    SLAB_CODE_MAP,
    SLAB_COLORS_LIGHT,
    SLAB_CONFIG,
    SLAB_ORDER,
    assign_slab,
    format_indian,
    get_next_slab,
    points_to_next,
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
# assign_slab (points-based)
# ---------------------------------------------------------------------------

class TestAssignSlab:
    """Tests for the assign_slab() function — points-based."""

    def test_zero_points(self) -> None:
        assert assign_slab(0) == "Unqualified"

    def test_just_below_slab_a(self) -> None:
        assert assign_slab(749) == "Unqualified"

    def test_at_slab_a(self) -> None:
        assert assign_slab(750) == "Slab A"

    def test_mid_slab_a(self) -> None:
        assert assign_slab(1500) == "Slab A"

    def test_slab_b(self) -> None:
        assert assign_slab(3000) == "Slab B"

    def test_slab_c(self) -> None:
        assert assign_slab(5000) == "Slab C"

    def test_slab_d(self) -> None:
        assert assign_slab(7000) == "Slab D"

    def test_slab_e(self) -> None:
        assert assign_slab(7500) == "Slab E"

    def test_slab_e_high(self) -> None:
        assert assign_slab(28000) == "Slab E"

    def test_nan_input(self) -> None:
        assert assign_slab(float("nan")) == "Unqualified"


# ---------------------------------------------------------------------------
# get_next_slab
# ---------------------------------------------------------------------------

class TestGetNextSlab:
    """Tests for the get_next_slab() function."""

    def test_unqualified(self) -> None:
        assert get_next_slab("Unqualified") == "Slab A"

    def test_slab_a(self) -> None:
        assert get_next_slab("Slab A") == "Slab B"

    def test_slab_d(self) -> None:
        assert get_next_slab("Slab D") == "Slab E"

    def test_last_slab(self) -> None:
        assert get_next_slab("Slab E") is None

    def test_unknown_slab(self) -> None:
        assert get_next_slab("Unknown") is None


# ---------------------------------------------------------------------------
# points_to_next
# ---------------------------------------------------------------------------

class TestPointsToNext:
    """Tests for the points_to_next() function."""

    def test_gap_exists(self) -> None:
        result = points_to_next(500, "Unqualified")
        assert result == 250.0

    def test_already_past_threshold(self) -> None:
        result = points_to_next(800, "Unqualified")
        assert result == 0.0

    def test_slab_a_to_b(self) -> None:
        result = points_to_next(2000, "Slab A")
        assert result == 1000.0

    def test_max_slab_returns_none(self) -> None:
        result = points_to_next(10000, "Slab E")
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
        required = {"slab", "slab_code", "range", "lower", "upper", "gift",
                     "gift_full", "category", "color", "color_light"}
        for cfg in SLAB_CONFIG:
            assert required.issubset(cfg.keys()), f"Missing keys in {cfg['slab']}"

    def test_color_light_values_are_hex(self) -> None:
        for cfg in SLAB_CONFIG:
            assert cfg["color_light"].startswith("#"), f"Invalid color_light in {cfg['slab']}"
            assert len(cfg["color_light"]) == 7, f"color_light should be #RRGGBB in {cfg['slab']}"

    def test_slab_colors_light_matches_config(self) -> None:
        for cfg in SLAB_CONFIG:
            assert SLAB_COLORS_LIGHT[cfg["slab"]] == cfg["color_light"]

    def test_slab_order_matches_config(self) -> None:
        assert SLAB_ORDER == [s["slab"] for s in SLAB_CONFIG]

    def test_last_slab_upper_is_inf(self) -> None:
        assert SLAB_CONFIG[-1]["upper"] == float("inf")

    def test_first_slab_starts_at_zero(self) -> None:
        assert SLAB_CONFIG[0]["lower"] == 0

    def test_slab_code_map_complete(self) -> None:
        for cfg in SLAB_CONFIG:
            assert cfg["slab_code"] in SLAB_CODE_MAP
            assert SLAB_CODE_MAP[cfg["slab_code"]] == cfg["slab"]


# ---------------------------------------------------------------------------
# SLAB_CODE_MAP (Excel code → slab name)
# ---------------------------------------------------------------------------

class TestSlabCodeMap:
    """Tests for the SLAB_CODE_MAP mapping."""

    def test_dash_maps_to_unqualified(self) -> None:
        assert SLAB_CODE_MAP["-"] == "Unqualified"

    def test_a_maps_to_slab_a(self) -> None:
        assert SLAB_CODE_MAP["A"] == "Slab A"

    def test_e_maps_to_slab_e(self) -> None:
        assert SLAB_CODE_MAP["E"] == "Slab E"
