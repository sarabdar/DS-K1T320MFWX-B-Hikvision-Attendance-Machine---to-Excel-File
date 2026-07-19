"""
attendance/__init__.py
======================
Public surface of the ``attendance`` package.
"""

from attendance.processor import build_attendance_dataframe

__all__ = ["build_attendance_dataframe"]
