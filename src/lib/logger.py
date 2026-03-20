"""Simple print-based logger for the dashboard application.

Provides info() and error() functions that print timestamped messages
to stdout. No external dependencies required.
"""

from datetime import datetime


def _timestamp() -> str:
    """Return current timestamp string in ISO format."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def info(msg: str) -> None:
    """Log an informational message to stdout.

    Args:
        msg: The message to log.
    """
    print(f"[INFO  {_timestamp()}] {msg}")


def error(msg: str) -> None:
    """Log an error message to stdout.

    Args:
        msg: The error message to log.
    """
    print(f"[ERROR {_timestamp()}] {msg}")
