"""Unit tests for helper functions defined in app.py."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from app import (
    NEXT_SLAB_MAP,
    POINTS_CONFIG,
    SLAB_CONFIG,
    SLAB_ORDER,
    assign_slab,
    calculate_qualified_points,
    calculate_qualified_volume,
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
# calculate_qualified_volume
# ---------------------------------------------------------------------------

class TestCalculateQualifiedVolume:
    """Tests for the calculate_qualified_volume() function."""

    def test_basic(self) -> None:
        assert calculate_qualified_volume(100, 50) == 150.0

    def test_zero_site(self) -> None:
        assert calculate_qualified_volume(200, 0) == 200.0

    def test_zero_shop(self) -> None:
        assert calculate_qualified_volume(0, 150) == 150.0

    def test_both_zero(self) -> None:
        assert calculate_qualified_volume(0, 0) == 0.0

    def test_nan_shop(self) -> None:
        assert calculate_qualified_volume(float("nan"), 100) == 100.0

    def test_nan_site(self) -> None:
        assert calculate_qualified_volume(50, float("nan")) == 50.0


# ---------------------------------------------------------------------------
# calculate_qualified_points
# ---------------------------------------------------------------------------

class TestCalculateQualifiedPoints:
    """Tests for the calculate_qualified_points() function."""

    def test_below_minimum(self) -> None:
        """Below 30 MT → 0 points."""
        assert calculate_qualified_points(29) == 0.0

    def test_at_minimum(self) -> None:
        """Exactly 30 MT → 30 * 25 = 750 (no milestone bonus)."""
        assert calculate_qualified_points(30) == 750

    def test_no_bonus(self) -> None:
        """40 MT → 40 * 25 = 1000 (below 50 MT, no bonus)."""
        assert calculate_qualified_points(40) == 1000

    def test_50mt_bonus(self) -> None:
        """60 MT → 60 * 25 * 1.10 = 1650."""
        assert calculate_qualified_points(60) == 1650

    def test_exactly_50mt(self) -> None:
        """50 MT → 50 * 25 * 1.10 = 1375."""
        assert calculate_qualified_points(50) == 1375

    def test_100mt_bonus(self) -> None:
        """100 MT → 100 * 25 * 1.20 = 3000."""
        assert calculate_qualified_points(100) == 3000

    def test_150mt_bonus(self) -> None:
        """150 MT → 150 * 25 * 1.30 = 4875."""
        assert calculate_qualified_points(150) == 4875

    def test_200mt_bonus(self) -> None:
        """200 MT → 200 * 25 * 1.50 = 7500."""
        assert calculate_qualified_points(200) == 7500

    def test_above_200mt(self) -> None:
        """250 MT → 250 * 25 * 1.50 = 9375."""
        assert calculate_qualified_points(250) == 9375

    def test_zero(self) -> None:
        assert calculate_qualified_points(0) == 0.0

    def test_nan(self) -> None:
        assert calculate_qualified_points(float("nan")) == 0.0


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

    def test_slab_b(self) -> None:
        assert assign_slab(3000) == "Slab B"

    def test_slab_c(self) -> None:
        assert assign_slab(5000) == "Slab C"

    def test_slab_d(self) -> None:
        assert assign_slab(7000) == "Slab D"

    def test_slab_e(self) -> None:
        assert assign_slab(7500) == "Slab E"

    def test_slab_e_high(self) -> None:
        assert assign_slab(15000) == "Slab E"

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
        required = {"slab", "range", "lower", "upper", "gift", "gift_full",
                     "category", "color"}
        for cfg in SLAB_CONFIG:
            assert required.issubset(cfg.keys()), f"Missing keys in {cfg['slab']}"

    def test_slab_order_matches_config(self) -> None:
        assert SLAB_ORDER == [s["slab"] for s in SLAB_CONFIG]

    def test_last_slab_upper_is_inf(self) -> None:
        assert SLAB_CONFIG[-1]["upper"] == float("inf")

    def test_first_slab_starts_at_zero(self) -> None:
        assert SLAB_CONFIG[0]["lower"] == 0


# ---------------------------------------------------------------------------
# End-to-end: volume → points → slab
# ---------------------------------------------------------------------------

class TestEndToEnd:
    """Integration-style tests for the full volume → points → slab pipeline."""

    def test_30mt_gets_slab_a(self) -> None:
        """30 MT → 750 pts → Slab A."""
        vol = calculate_qualified_volume(30, 0)
        pts = calculate_qualified_points(vol)
        slab = assign_slab(pts)
        assert pts == 750
        assert slab == "Slab A"

    def test_100mt_gets_slab_b(self) -> None:
        """100 MT → 3000 pts → Slab B."""
        vol = calculate_qualified_volume(60, 40)
        pts = calculate_qualified_points(vol)
        slab = assign_slab(pts)
        assert pts == 3000
        assert slab == "Slab B"

    def test_200mt_gets_slab_e(self) -> None:
        """200 MT → 7500 pts → Slab E."""
        vol = calculate_qualified_volume(150, 50)
        pts = calculate_qualified_points(vol)
        slab = assign_slab(pts)
        assert pts == 7500
        assert slab == "Slab E"

    def test_below_minimum_unqualified(self) -> None:
        """20 MT → 0 pts → Unqualified."""
        vol = calculate_qualified_volume(15, 5)
        pts = calculate_qualified_points(vol)
        slab = assign_slab(pts)
        assert pts == 0
        assert slab == "Unqualified"
