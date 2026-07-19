"""
config.py
=========
Central configuration for the Hikvision Attendance tool.
Loads all settings from the .env file and exposes them as
typed constants consumed by other modules.
"""

import os
from dotenv import load_dotenv

load_dotenv(override=True)

# ── Device ────────────────────────────────────────────────
DEVICE_IP: str = os.getenv("DEVICE_IP", "")
USERNAME: str = os.getenv("USERNAME", "")
PASSWORD: str = os.getenv("PASSWORD", "")

# ── API ───────────────────────────────────────────────────
API_URL: str = (
    f"http://{DEVICE_IP}"
    f"/ISAPI/AccessControl/AcsEvent?format=json"
)

HTTP_HEADERS: dict = {
    "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64)",
    "Content-Type": "application/json",
}

# ── Fetch settings ────────────────────────────────────────
PAGE_SIZE: int = 500
REQUEST_TIMEOUT: int = 15          # seconds
INTER_PAGE_DELAY: float = 0.1      # seconds between pages

# ── Excel ─────────────────────────────────────────────────
EXCEL_SHEET_NAME: str = "Attendance Sheet"
