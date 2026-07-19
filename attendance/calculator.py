"""
attendance/calculator.py
========================
Pure functions for computing shift durations from raw datetime
objects. No I/O, no pandas — easy to unit-test in isolation.
"""

from datetime import datetime

import pandas as pd


def calc_shift_minutes(check_in: datetime, check_out: datetime) -> int:
    """
    Return the duration of a shift in whole minutes.

    Handles the edge case where check-out is on the *next* calendar
    day (i.e. the employee worked a night shift that crossed midnight).

    Parameters
    ----------
    check_in:
        Earliest check-in timestamp for the day.
    check_out:
        Latest check-out timestamp for the day.

    Returns
    -------
    int
        Duration in minutes (always ≥ 0).
    """
    if check_out < check_in:
        check_out = check_out + pd.Timedelta(days=1)

    delta = check_out - check_in
    return int(delta.total_seconds() / 60)


def minutes_to_hm(total_minutes: int) -> str:
    """
    Format a minute count as a human-readable ``"Xh Ym"`` string.

    Parameters
    ----------
    total_minutes:
        Non-negative integer number of minutes.

    Returns
    -------
    str
        E.g. ``"8h 30m"`` or ``"0h 0m"``.
    """
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours}h {minutes}m"


def format_time(dt: datetime) -> str:
    """
    Format a datetime into a compact 12-hour clock string.

    Leading zeroes are stripped and AM/PM is lower-cased so the
    output matches the original sheet style (e.g. ``"9:05am"``).

    Parameters
    ----------
    dt:
        A datetime (or pandas Timestamp) to format.

    Returns
    -------
    str
        E.g. ``"9:05am"``, ``"12:00pm"``.
    """
    return dt.strftime("%I:%M%p").lstrip("0").lower()
