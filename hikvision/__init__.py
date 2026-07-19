"""
hikvision/__init__.py
=====================
Public surface of the ``hikvision`` package.

Consumers only need to import from here; internal module
layout can change freely without breaking callers.
"""

from hikvision.client import build_session
from hikvision.fetcher import fetch_events

__all__ = ["build_session", "fetch_events"]
