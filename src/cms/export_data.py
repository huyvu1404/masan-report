"""
Export-based data fetch via CMS export API.

Flow per date chunk:
  1. count_buzzes()      – trackingTotalBuzzes to get row count
  2. _split_ranges()     – recursively halve range until each chunk ≤ MAX_EXPORT_ROWS
  3. trigger_export()    – exportBuzzes mutation → fileName
  4. poll_export()       – poll exportManagementByOwner → downloadLink
  5. download_and_read() – GET + decompress ZIP + read Excel → DataFrame

The exported Excel is already in the format expected by process_data(),
so no additional column mapping is needed.

Entry point: fetch_data_export()
"""

from __future__ import annotations

import io
import time
import zipfile
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd
import requests

from .fetch_data import (
    _CMS_URL,
    _BASE_HEADERS,
    _ALL_TYPES,
    _ALL_SENTIMENTS,
    _MAX_RETRIES,
    _RETRY_BACKOFF,
    _RETRIABLE,
)

MAX_EXPORT_ROWS = 500_000
POLL_MAX_WAIT   = 900   # 15 min
POLL_INIT_SEC   = 10
POLL_MAX_SEC    = 30

_COUNT_QUERY = """
query countBuzzes($input: IndexesInput!, $filter: FilterBuzzInput) {
  trackingTotalBuzzes(input: $input, filter: $filter) {
    total
  }
}
"""

_EXPORT_MUTATION = """
mutation exportBuzzes(
  $input: IndexesInfoInput!,
  $filter: FilterBuzzInput,
  $type: ExportBuzzFileType!,
  $projectFieldsDetail: ProjectFieldsDetail!,
  $fieldExportType: FieldExportTypeEnum!,
  $select: ExportBuzzFileSelect!
) {
  exportBuzzes(
    input: $input
    filter: $filter
    projectFieldsDetail: $projectFieldsDetail
    fieldExportType: $fieldExportType
    type: $type
    select: $select
  ) {
    status
    message
    data {
      status
      fileName
    }
  }
}
"""

_STATUS_QUERY = """
query exportManagementByOwner {
  exportManagementByOwner {
    status
    message
    data {
      status
      taskName
      downloadLink
      createdDate
      fileName
    }
  }
}
"""

_EXPORT_SELECT = {
    "title": True, "content": True, "description": True,
    "urlComment": True, "urlTopic": True, "topic": True,
    "publishedDate": True, "sentiment": True, "level": True,
    "siteName": True, "siteId": True, "parentId": True,
    "labelsName": True, "type": True, "profile": True,
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _gql(headers: dict, payload: dict, timeout: int = 60) -> dict:
    """POST GraphQL with retry on timeout / connection error."""
    last_exc: Exception = RuntimeError("no attempts")
    for attempt in range(_MAX_RETRIES + 1):
        try:
            resp = requests.post(_CMS_URL, json=payload, headers=headers, timeout=timeout)
            resp.raise_for_status()
            body = resp.json()
            if body.get("errors"):
                raise ValueError(f"GraphQL errors: {body['errors']}")
            return body.get("data", {})
        except _RETRIABLE as exc:
            last_exc = exc
            if attempt < _MAX_RETRIES:
                wait = _RETRY_BACKOFF[attempt]
                print(f"[export] {type(exc).__name__} attempt {attempt + 1}/{_MAX_RETRIES + 1}, retry in {wait}s")
                time.sleep(wait)
    raise last_exc


# ── Step 1: count ─────────────────────────────────────────────────────────────

def count_buzzes(
    headers: dict,
    topic_ids: list[str],
    from_date: str,
    to_date: str,
    types: Optional[list[str]] = None,
    sentiments: Optional[list[str]] = None,
) -> int:
    """Return total buzz count for the date range using the published-date filter."""
    data = _gql(headers, {
        "operationName": "countBuzzes",
        "variables": {
            "input": {"indexes": topic_ids},
            "filter": {
                "publishedFromDate": from_date,
                "publishedToDate":   to_date,
                "types":             types or _ALL_TYPES,
                "sentiments":        sentiments or _ALL_SENTIMENTS,
                "isDeleted":         False,
                "isExactQuery":      False,
            },
        },
        "query": _COUNT_QUERY,
    })
    return data.get("trackingTotalBuzzes", {}).get("total", 0)


# ── Step 2: split ─────────────────────────────────────────────────────────────

def _split_ranges(
    headers: dict,
    topic_ids: list[str],
    from_date: str,
    to_date: str,
    types: list[str],
    sentiments: list[str],
    threshold: int,
) -> list[tuple[str, str]]:
    """Recursively halve [from_date, to_date] until each chunk has ≤ threshold rows."""
    total = count_buzzes(headers, topic_ids, from_date, to_date, types, sentiments)
    print(f"[export] count {from_date[:10]} → {to_date[:10]}: {total:,}")

    if total <= threshold:
        return [(from_date, to_date)]

    start = datetime.strptime(from_date[:10], "%Y-%m-%d")
    end   = datetime.strptime(to_date[:10],   "%Y-%m-%d")

    if start >= end:
        # Single day, cannot split further — proceed anyway
        return [(from_date, to_date)]

    mid        = start + (end - start) // 2
    mid_end    = f"{mid.strftime('%Y-%m-%d')} 23:59:59"
    next_start = f"{(mid + timedelta(days=1)).strftime('%Y-%m-%d')} 00:00:00"

    left  = _split_ranges(headers, topic_ids, from_date,   mid_end,    types, sentiments, threshold)
    right = _split_ranges(headers, topic_ids, next_start,  to_date,    types, sentiments, threshold)
    return left + right


# ── Step 3: trigger export ────────────────────────────────────────────────────

def trigger_export(
    headers: dict,
    topic_id_name_pairs: list[dict],
    from_date: str,
    to_date: str,
    project_info: dict,
    types: Optional[list[str]] = None,
    sentiments: Optional[list[str]] = None,
) -> str:
    """Fire exportBuzzes mutation; return fileName used for polling."""
    data = _gql(headers, {
        "operationName": "exportBuzzes",
        "variables": {
            "input": {"indexes": topic_id_name_pairs},
            "filter": {
                "publishedFromDate": from_date,
                "publishedToDate":   to_date,
                "types":             types or _ALL_TYPES,
                "sentiments":        sentiments or _ALL_SENTIMENTS,
                "polarity":          ["NONE", "EQUAL", "LESS", "GREATER"],
                "riskGroups":        [],
                "isDeleted":         False,
                "isExactQuery":      False,
            },
            "type": "EXCEL",
            "projectFieldsDetail": {
                "labels":       project_info.get("labels", []),
                "crisisTags":   [],
                "campaignTags": [],
                "filters":      project_info.get("filters", []),
            },
            "fieldExportType": "NESTED",
            "select": _EXPORT_SELECT,
        },
        "query": _EXPORT_MUTATION,
    })

    file_name = (data.get("exportBuzzes") or {}).get("data", {}).get("fileName")
    if not file_name:
        raise ValueError(f"No fileName in exportBuzzes response: {data}")
    print(f"[export] triggered → {file_name}")
    return file_name


# ── Step 4: poll ──────────────────────────────────────────────────────────────

def poll_export(
    headers: dict,
    file_name: str,
    max_wait: int = POLL_MAX_WAIT,
) -> dict:
    """
    Poll exportManagementByOwner until the job matching file_name is done.

    Returns the export record (contains downloadLink).
    Raises RuntimeError on job failure, TimeoutError when max_wait exceeded.
    """
    deadline = time.time() + max_wait
    interval = POLL_INIT_SEC

    while time.time() < deadline:
        time.sleep(interval)
        interval = min(interval + 5, POLL_MAX_SEC)

        try:
            data = _gql(headers, {
                "operationName": "exportManagementByOwner",
                "variables": {},
                "query": _STATUS_QUERY,
            }, timeout=30)
        except Exception as exc:
            print(f"[export] poll error (will retry): {exc}")
            continue

        exports: list[dict] = (data.get("exportManagementByOwner") or {}).get("data") or []

        for record in exports:
            if record.get("taskName") != file_name:
                continue
            status = record.get("status")
            print(f"[export] status={status}  ({file_name})")
            if status == "done":
                return record
            if status == "fail":
                raise RuntimeError(f"Export job failed: {file_name}")

    raise TimeoutError(f"Export {file_name} did not complete within {max_wait}s")


# ── Step 5: download + decompress ─────────────────────────────────────────────

def download_and_read(headers: dict, download_link: str) -> pd.DataFrame:
    """GET the export file, decompress ZIP if needed, read Excel/CSV → DataFrame."""
    last_exc: Exception = RuntimeError("no download attempts")
    for attempt in range(_MAX_RETRIES + 1):
        try:
            resp = requests.get(download_link, headers=headers, timeout=120)
            resp.raise_for_status()
            content = resp.content
            break
        except _RETRIABLE as exc:
            last_exc = exc
            if attempt < _MAX_RETRIES:
                wait = _RETRY_BACKOFF[attempt]
                print(f"[export] download timeout, retry in {wait}s")
                time.sleep(wait)
    else:
        raise last_exc

    if zipfile.is_zipfile(io.BytesIO(content)):
        with zipfile.ZipFile(io.BytesIO(content)) as zf:
            names = [n for n in zf.namelist() if n.lower().endswith((".xlsx", ".csv"))]
            if not names:
                raise ValueError(f"No xlsx/csv in ZIP. Contents: {zf.namelist()}")
            frames = []
            for name in names:
                raw = zf.read(name)
                df  = (pd.read_excel(io.BytesIO(raw))
                       if name.lower().endswith(".xlsx")
                       else pd.read_csv(io.BytesIO(raw)))
                frames.append(df)
            return pd.concat(frames, ignore_index=True)

    # Direct file (not zipped)
    try:
        return pd.read_excel(io.BytesIO(content))
    except Exception:
        return pd.read_csv(io.BytesIO(content))


# ── Entry point ───────────────────────────────────────────────────────────────

def fetch_data_export(
    access_token: str,
    refresh_token: str,
    topic_ids: list[str],
    topic_id_to_name: dict[str, str],
    from_date: str,
    to_date: str,
    project_info: dict,
    types: Optional[list[str]] = None,
    sentiments: Optional[list[str]] = None,
) -> pd.DataFrame:
    """
    Fetch CMS data via export API.

    Splits the date range when total rows > MAX_EXPORT_ROWS.
    Retries failed export jobs up to _MAX_RETRIES times.

    Returns a DataFrame ready for process_data().
    """
    headers = {
        **_BASE_HEADERS,
        "x-token":         f"Bearer {access_token}",
        "x-refresh-token": f"Bearer {refresh_token}",
    }
    resolved_types      = types or _ALL_TYPES
    resolved_sentiments = sentiments or _ALL_SENTIMENTS

    topic_id_name_pairs = [
        {"id": tid, "name": topic_id_to_name.get(tid, tid)}
        for tid in topic_ids
    ]

    # Determine chunks
    chunks = _split_ranges(
        headers, topic_ids, from_date, to_date,
        resolved_types, resolved_sentiments, MAX_EXPORT_ROWS,
    )
    print(f"[export] {len(chunks)} chunk(s): {[f'{a[:10]}→{b[:10]}' for a, b in chunks]}")

    # Export each chunk (with per-chunk retry on job failure)
    frames: list[pd.DataFrame] = []
    for chunk_from, chunk_to in chunks:
        for attempt in range(_MAX_RETRIES + 1):
            try:
                file_name   = trigger_export(headers, topic_id_name_pairs, chunk_from, chunk_to, project_info, resolved_types, resolved_sentiments)
                record      = poll_export(headers, file_name)
                raw_df      = download_and_read(headers, record["downloadLink"])
                frames.append(raw_df)
                print(f"[export] chunk {chunk_from[:10]}→{chunk_to[:10]}: {len(raw_df):,} rows")
                break
            except (RuntimeError, TimeoutError) as exc:
                if attempt < _MAX_RETRIES:
                    wait = _RETRY_BACKOFF[attempt]
                    print(f"[export] chunk {chunk_from[:10]}→{chunk_to[:10]} failed ({exc}), retry in {wait}s")
                    time.sleep(wait)
                else:
                    raise

    if not frames:
        return pd.DataFrame()

    raw = pd.concat(frames, ignore_index=True)
    print(f"[export] total rows: {len(raw):,} | columns: {list(raw.columns)}")
    return raw
