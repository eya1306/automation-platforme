"""Runtime configuration.

Every value can be overridden with an environment variable, so the same code
runs on a laptop and on a shared server without edits.
"""

from __future__ import annotations

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

# Where uploads and generated reports live. One folder per run.
WORKSPACE = Path(os.environ.get("AUTOMATION_WORKSPACE", BASE_DIR / "workspace"))
JOBS_DIR = WORKSPACE / "jobs"

# Largest request the server will accept, uploads included.
MAX_UPLOAD_MB = int(os.environ.get("AUTOMATION_MAX_UPLOAD_MB", "250"))

# How many runs execute at once. The Excel and Word engines both shell out to
# LibreOffice for legacy formats, which is memory-hungry, so keep this small.
MAX_WORKERS = int(os.environ.get("AUTOMATION_MAX_WORKERS", "2"))

# Finished runs older than this are deleted, with their files, on next sweep.
RETENTION_HOURS = int(os.environ.get("AUTOMATION_RETENTION_HOURS", "24"))

# How many runs stay listed in the history rail.
HISTORY_LIMIT = int(os.environ.get("AUTOMATION_HISTORY_LIMIT", "40"))

# --------------------------------------------------------------------------- #
# Branding
# --------------------------------------------------------------------------- #

STATIC_DIR = Path(__file__).resolve().parent / "static"

# Shown in the masthead, the browser tab and the start-up banner.
APP_NAME = os.environ.get("AUTOMATION_APP_NAME", "AutoGeneration Platform")
APP_TAGLINE = os.environ.get(
    "AUTOMATION_APP_TAGLINE", "Document and spreadsheet change review"
)

# The organisation the platform belongs to. Its logo is drawn on the left of
# the masthead: drop the official file into platform_app/static/ under one of
# the names below and it is picked up on the next page load. With no file
# present the masthead falls back to the organisation name set in text, so the
# page never shows a broken image.
ORG_NAME = os.environ.get("AUTOMATION_ORG_NAME", "Capgemini Engineering")
LOGO_NAMES = ("logo.svg", "logo.png", "logo.webp", "logo.jpg", "logo.jpeg")


def logo_filename() -> str | None:
    """Name of the logo file inside static/, or None if none was supplied."""
    for name in LOGO_NAMES:
        if (STATIC_DIR / name).is_file():
            return name
    return None


HOST = os.environ.get("AUTOMATION_HOST", "127.0.0.1")
PORT = int(os.environ.get("AUTOMATION_PORT", "8000"))


def ensure_dirs() -> None:
    JOBS_DIR.mkdir(parents=True, exist_ok=True)
