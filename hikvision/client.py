"""
hikvision/client.py
===================
Creates and returns a pre-configured requests.Session
that talks to the Hikvision device.

All network-level concerns (auth, proxies, headers)
are isolated here so callers never touch raw HTTP.
"""

import requests
from requests.auth import HTTPDigestAuth
from config import USERNAME, PASSWORD, HTTP_HEADERS


def build_session() -> requests.Session:
    """
    Return a requests.Session pre-configured with:
    - Digest authentication for the Hikvision device.
    - No proxy settings (prevents corporate proxy interference).
    - Default JSON headers.

    Returns
    -------
    requests.Session
    """
    session = requests.Session()

    session.auth = HTTPDigestAuth(USERNAME, PASSWORD)

    # Explicitly bypass any system proxy settings.
    session.proxies = {"http": None, "https": None}

    session.headers.update(HTTP_HEADERS)

    return session
