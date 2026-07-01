

from __future__ import annotations

import io
import json
import os
import re

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

_URL_PATTERNS = [
    re.compile(r"/file/d/([a-zA-Z0-9_-]+)"),   # /file/d/<ID>/view
    re.compile(r"[?&]id=([a-zA-Z0-9_-]+)"),     # ?id=<ID>
]


def extract_file_id(file_id_or_url: str) -> str:
    """Return the bare Drive file ID from either a URL or a bare ID."""
    s = file_id_or_url.strip()
    for pat in _URL_PATTERNS:
        m = pat.search(s)
        if m:
            return m.group(1)
    return s  # assume it's already a bare ID


def build_service():
    """
    Build an authenticated Drive v3 service.

    Args:
        sa_json: Service-account credentials as a JSON string.
                 Falls back to env var GOOGLE_SERVICE_ACCOUNT_JSON.
        sa_file: Path to a service-account credentials JSON file.
                 Falls back to env var GOOGLE_SERVICE_ACCOUNT_FILE.
    """
    sa_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    sa_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")

    if sa_json:
        info  = json.loads(sa_json)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    elif sa_file:
        creds = service_account.Credentials.from_service_account_file(sa_file, scopes=SCOPES)
    else:
        raise ValueError(
            "No Google credentials found. "
            "Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_FILE."
        )

    return build("drive", "v3", credentials=creds, cache_discovery=False)


def download_file(service, file_id_or_url: str) -> bytes:
    """
    Download a Drive file by ID (or URL) and return its raw bytes.

    For Google Workspace files (Sheets/Docs) this automatically exports to xlsx.
    For regular uploads (.xlsx, .xls) it downloads directly.
    """
    fid = extract_file_id(file_id_or_url)

    # Check mime type to decide between export and direct download
    meta = service.files().get(fileId=fid, fields="mimeType,name", supportsAllDrives=True ).execute()
    mime = meta.get("mimeType", "")

    XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    if mime == "application/vnd.google-apps.spreadsheet":
        # Google Sheets → export as xlsx
        request = service.files().export_media(fileId=fid, mimeType=XLSX_MIME)
    else:
        # Regular file (xlsx / xls / csv) → direct download
        request = service.files().get_media(fileId=fid)

    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request, chunksize=8 * 1024 * 1024)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    return buf.getvalue()
