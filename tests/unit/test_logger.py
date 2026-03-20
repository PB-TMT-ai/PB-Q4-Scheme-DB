"""Unit tests for src/lib/logger module."""

from __future__ import annotations

import sys
from io import StringIO
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from src.lib.logger import error, info


class TestLogger:
    """Tests for the logger module."""

    def test_info_output(self, capsys: object) -> None:
        info("test message")
        captured = capsys.readouterr()
        assert "[INFO" in captured.out
        assert "test message" in captured.out

    def test_error_output(self, capsys: object) -> None:
        error("error message")
        captured = capsys.readouterr()
        assert "[ERROR" in captured.out
        assert "error message" in captured.out
