"""
hikvision/fetcher.py
====================
Responsible for paginating through the Hikvision ACS event API
and returning only valid attendance log entries.

Keeps all pagination + retry logic in one place so the rest of
the codebase never has to worry about HTTP or device quirks.
"""

import json
import logging
from time import sleep
from typing import Any

import requests
from requests.auth import HTTPDigestAuth

from config import API_URL, PAGE_SIZE, REQUEST_TIMEOUT, INTER_PAGE_DELAY, USERNAME, PASSWORD

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def fetch_events(
    session: requests.Session,
    start_time: str,
    end_time: str,
) -> list[dict[str, Any]]:
    """
    Fetch all attendance events from the Hikvision device for the
    given time window.

    The device returns results in pages of ``PAGE_SIZE`` records.
    This function handles pagination transparently and filters the
    raw response to keep only genuine check-in / check-out records.

    Parameters
    ----------
    session:
        A pre-authenticated requests.Session (from ``hikvision.client``).
    start_time:
        ISO-8601 string, e.g. ``"2025-06-01T00:00:00+05:00"``.
    end_time:
        ISO-8601 string, e.g. ``"2025-06-30T23:59:59+05:00"``.

    Returns
    -------
    list[dict]
        Filtered list of raw event dicts that contain
        ``employeeNoString``, ``time``, and ``attendanceStatus``.
    """
    raw_events: list[dict[str, Any]] = []
    position = 0

    print(f"Starting data extraction from Hikvision device…")

    while True:
        payload = _build_payload(position, start_time, end_time)

        try:
            response = session.post(
                API_URL,
                auth=HTTPDigestAuth(USERNAME, PASSWORD),
                data=json.dumps(payload),
                timeout=REQUEST_TIMEOUT,
            )
            response.raise_for_status()

        except requests.RequestException as exc:
            logger.error("Network error during pagination at position %d: %s", position, exc)
            print(f"\nError occurred during pagination loop: {exc}")
            break

        page_events, status_strg, total_matches = _parse_response(response)

        if not page_events:
            break

        print(
            f"Fetched records {position} to "
            f"{position + len(page_events)} "
            f"(Total Available: {total_matches})"
        )

        raw_events.extend(_filter_attendance(page_events))

        if status_strg != "MORE":
            break

        position += len(page_events)
        sleep(INTER_PAGE_DELAY)

    return raw_events


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _build_payload(position: int, start_time: str, end_time: str) -> dict:
    """Construct the ACS event search request body."""
    return {
        "AcsEventCond": {
            "searchID": "1",
            "searchResultPosition": position,
            "maxResults": PAGE_SIZE,
            "startTime": start_time,
            "endTime": end_time,
            "major": 0,
            "minor": 0,
        }
    }


def _parse_response(
    response: requests.Response,
) -> tuple[list[dict], str, int]:
    """
    Extract the event list, pagination status, and total match count
    from a raw API response.

    Returns
    -------
    (info_list, status_strg, total_matches)
    """
    data = response.json()
    acs_event = data.get("AcsEvent", {})

    info_list: list[dict] = acs_event.get("InfoList", [])
    status_strg: str = acs_event.get("responseStatusStrg", "OK")
    total_matches: int = acs_event.get("totalMatches", 0)

    return info_list, status_strg, total_matches


def _filter_attendance(events: list[dict]) -> list[dict]:
    """
    Keep only events that represent a real attendance action:
    must have ``employeeNoString``, a ``time``, and an
    ``attendanceStatus`` field.
    """
    return [
        log for log in events
        if (
            "employeeNoString" in log
            and log.get("time")
            and log.get("attendanceStatus")
        )
    ]
