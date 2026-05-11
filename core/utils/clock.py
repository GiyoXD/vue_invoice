"""
clock.py — Centralized time helper for the project.

All timestamps in this app use Asia/Phnom Penh (ICT = UTC+7).
Import from here so datetime/timezone/timedelta never scatter across modules.

Usage:
    from core.utils.clock import now, timestamp

    dt = now()                    # datetime object, tz-aware, ICT
    ts = timestamp()              # ISO 8601 string  e.g. "2026-05-09T16:27:44+07:00"
    ts = timestamp("%d/%m/%Y")    # custom format    e.g. "09/05/2026"
"""

from datetime import datetime, timezone, timedelta

# Phnom Penh / Indochina Time — fixed UTC+7, no DST
ICT = timezone(timedelta(hours=7))


def now() -> datetime:
    """Return current datetime, timezone-aware, in Phnom Penh / ICT (UTC+7)."""
    return datetime.now(ICT)


def timestamp(fmt: str | None = None) -> str:
    """
    Return current time as a string.

    Args:
        fmt: strftime format string. Defaults to ISO 8601.

    Returns:
        Formatted time string in ICT.

    Examples:
        timestamp()             → "2026-05-09T16:27:44+07:00"
        timestamp("%d/%m/%Y")   → "09/05/2026"
        timestamp("%H:%M")      → "16:27"
    """
    dt = now()
    if fmt:
        return dt.strftime(fmt)
    return dt.isoformat()
