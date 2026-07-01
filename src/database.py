"""
SQLite helper — persists report metadata and pipeline task state.

Tables
------
reports
  id          INTEGER PRIMARY KEY AUTOINCREMENT
  main_brand  TEXT
  competitors TEXT  (JSON array, sorted)
  start_date  TEXT
  end_date    TEXT
  report_name TEXT
  created_at  TEXT

tasks
  task_id     TEXT PRIMARY KEY
  status      TEXT  (pending | running | done | failed)
  step        TEXT
  file_path   TEXT
  filename    TEXT
  error       TEXT
  main_brand  TEXT
  competitors TEXT  (JSON array)
  start_date  TEXT
  end_date    TEXT
  created_at  TEXT
"""

from __future__ import annotations

import json
import os
import sqlite3
import threading
from datetime import datetime
from pathlib import Path

DB_PATH = Path(os.getenv("DB_PATH", "reports.db"))
_lock = threading.Lock()

_DDL = """
CREATE TABLE IF NOT EXISTS reports (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    main_brand  TEXT NOT NULL,
    competitors TEXT NOT NULL,
    start_date  TEXT NOT NULL,
    end_date    TEXT NOT NULL,
    report_name TEXT NOT NULL,
    created_at  TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS tasks (
    task_id          TEXT PRIMARY KEY,
    status           TEXT NOT NULL DEFAULT 'pending',
    step             TEXT,
    file_path        TEXT,
    filename         TEXT,
    error            TEXT,
    main_brand       TEXT,
    competitors      TEXT,
    start_date       TEXT,
    end_date         TEXT,
    created_at       TEXT NOT NULL,
    duration_seconds REAL
);
"""


def _connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


_MIGRATIONS = [
    "ALTER TABLE tasks ADD COLUMN duration_seconds REAL",
]


def init_db() -> None:
    with _lock:
        conn = _connect()
        conn.executescript(_DDL)
        for sql in _MIGRATIONS:
            try:
                conn.execute(sql)
            except sqlite3.OperationalError:
                pass  # column already exists
        conn.commit()
        conn.close()


# ── Reports ───────────────────────────────────────────────────────────────────

def find_report(
    main_brand: str,
    competitors: list[str],
    start_date: str,
    end_date: str,
) -> dict | None:
    competitors_json = json.dumps(sorted(competitors))
    with _lock:
        conn = _connect()
        row = conn.execute(
            "SELECT * FROM reports "
            "WHERE main_brand=? AND competitors=? AND start_date=? AND end_date=? "
            "ORDER BY id DESC LIMIT 1",
            (main_brand, competitors_json, start_date, end_date),
        ).fetchone()
        conn.close()
    return dict(row) if row else None


def save_report(
    main_brand: str,
    competitors: list[str],
    start_date: str,
    end_date: str,
    report_name: str,
) -> None:
    competitors_json = json.dumps(sorted(competitors))
    created_at = datetime.now().isoformat()
    with _lock:
        conn = _connect()
        conn.execute(
            "INSERT INTO reports "
            "(main_brand, competitors, start_date, end_date, report_name, created_at) "
            "VALUES (?,?,?,?,?,?)",
            (main_brand, competitors_json, start_date, end_date, report_name, created_at),
        )
        conn.commit()
        conn.close()


def list_reports() -> list[dict]:
    with _lock:
        conn = _connect()
        rows = conn.execute("SELECT * FROM reports ORDER BY id DESC").fetchall()
        conn.close()
    result = []
    for row in rows:
        r = dict(row)
        r["competitors"] = json.loads(r["competitors"])
        result.append(r)
    return result


def delete_report_by_name(report_name: str) -> int:
    with _lock:
        conn = _connect()
        cursor = conn.execute("DELETE FROM reports WHERE report_name=?", (report_name,))
        conn.commit()
        deleted = cursor.rowcount
        conn.close()
    return deleted


# ── Tasks ─────────────────────────────────────────────────────────────────────

def _task_row_to_dict(row: sqlite3.Row) -> dict:
    r = dict(row)
    if r.get("competitors"):
        r["competitors"] = json.loads(r["competitors"])
    else:
        r["competitors"] = []
    return r


def claim_task(
    task_id: str,
    main_brand: str,
    competitors: list[str],
    start_date: str,
    end_date: str,
) -> tuple[str, bool]:
    """
    Atomically check for an active task then create one if none exists.

    Returns (task_id, created):
      - created=True  → new task inserted, caller should start the pipeline
      - created=False → existing pending/running task found, caller should reuse it
    """
    competitors_json = json.dumps(competitors)
    created_at = datetime.now().isoformat()
    with _lock:
        conn = _connect()
        row = conn.execute(
            "SELECT task_id FROM tasks "
            "WHERE main_brand=? AND competitors=? AND start_date=? AND end_date=? "
            "  AND status IN ('pending','running') "
            "ORDER BY created_at DESC LIMIT 1",
            (main_brand, competitors_json, start_date, end_date),
        ).fetchone()
        if row:
            conn.close()
            return row["task_id"], False
        conn.execute(
            "INSERT INTO tasks "
            "(task_id, status, step, file_path, filename, error, "
            " main_brand, competitors, start_date, end_date, created_at) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (task_id, "pending", None, None, None, None,
             main_brand, competitors_json, start_date, end_date, created_at),
        )
        conn.commit()
        conn.close()
    return task_id, True


def update_task(task_id: str, **kwargs) -> None:
    if not kwargs:
        return
    allowed = {"status", "step", "file_path", "filename", "error"}
    fields = {k: v for k, v in kwargs.items() if k in allowed}
    if not fields:
        return

    now = datetime.now()

    with _lock:
        conn = _connect()

        if fields.get("status") in ("done", "failed"):
            row = conn.execute(
                "SELECT created_at FROM tasks WHERE task_id=?", (task_id,)
            ).fetchone()
            if row and row["created_at"]:
                created = datetime.fromisoformat(row["created_at"])
                fields["duration_seconds"] = round((now - created).total_seconds(), 1)

        set_clause = ", ".join(f"{k}=?" for k in fields)
        values = list(fields.values()) + [task_id]
        conn.execute(f"UPDATE tasks SET {set_clause} WHERE task_id=?", values)
        conn.commit()
        conn.close()


def get_task(task_id: str) -> dict | None:
    with _lock:
        conn = _connect()
        row = conn.execute("SELECT * FROM tasks WHERE task_id=?", (task_id,)).fetchone()
        conn.close()
    return _task_row_to_dict(row) if row else None


def list_tasks() -> dict[str, dict]:
    with _lock:
        conn = _connect()
        rows = conn.execute("SELECT * FROM tasks ORDER BY created_at DESC").fetchall()
        conn.close()
    return {row["task_id"]: _task_row_to_dict(row) for row in rows}
