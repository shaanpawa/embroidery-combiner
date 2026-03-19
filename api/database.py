"""
SQLite persistence layer for Micro Automation — Combo Builder.
Replaces the in-memory sessions dict in server.py.

DB file: {project_root}/data/micro.db
Session files: {project_root}/data/sessions/{sid}/dst/ and output/
"""

import json
import os
import sqlite3
import threading
from datetime import datetime, timezone
from typing import Dict, List, Optional

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(PROJECT_ROOT, "data")
DB_PATH = os.path.join(DATA_DIR, "micro.db")
SESSIONS_DIR = os.path.join(DATA_DIR, "sessions")

# ---------------------------------------------------------------------------
# Connection management (thread-safe singleton)
# ---------------------------------------------------------------------------

_db_lock = threading.Lock()
_connection: Optional[sqlite3.Connection] = None


def get_db() -> sqlite3.Connection:
    """Return a shared SQLite connection, creating tables on first call."""
    global _connection
    if _connection is not None:
        return _connection
    with _db_lock:
        if _connection is not None:
            return _connection
        _connection = init_db()
        return _connection


def init_db() -> sqlite3.Connection:
    """Create the database file and tables if they don't exist."""
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            id               TEXT PRIMARY KEY,
            user_email       TEXT DEFAULT 'local',
            name             TEXT NOT NULL,
            created_at       TEXT NOT NULL,
            updated_at       TEXT NOT NULL,
            excel_filename   TEXT,
            entries_json     TEXT,
            groups_json      TEXT,
            combos_json      TEXT,
            dst_programs_json TEXT,
            gap_mm           REAL DEFAULT 3.0,
            column_gap_mm    REAL DEFAULT 5.0,
            exported         INTEGER DEFAULT 0,
            exported_at      TEXT
        )
    """)
    conn.commit()
    return conn


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _now() -> str:
    """ISO-8601 timestamp in UTC."""
    return datetime.now(timezone.utc).isoformat()


def _row_to_dict(row: Optional[sqlite3.Row]) -> Optional[dict]:
    """Convert a sqlite3.Row to a plain dict, deserialising JSON fields."""
    if row is None:
        return None
    d = dict(row)
    for key in ("entries_json", "groups_json", "combos_json", "dst_programs_json"):
        raw = d.get(key)
        d[key] = json.loads(raw) if raw else []
    d["exported"] = bool(d.get("exported", 0))
    return d


def _json_len(raw: Optional[str]) -> int:
    """Return the length of a JSON array stored as text, without full parse."""
    if not raw:
        return 0
    try:
        return len(json.loads(raw))
    except (json.JSONDecodeError, TypeError):
        return 0


# ---------------------------------------------------------------------------
# CRUD functions
# ---------------------------------------------------------------------------

def create_session(sid: str, name: str, user_email: str = "local") -> Dict:
    """Insert a new session row and return it as a dict."""
    now = _now()
    db = get_db()
    db.execute(
        """
        INSERT INTO sessions (id, user_email, name, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?)
        """,
        (sid, user_email, name, now, now),
    )
    db.commit()
    return get_session(sid)


def get_session(sid: str) -> Optional[Dict]:
    """Return a full session dict or None if not found."""
    db = get_db()
    row = db.execute("SELECT * FROM sessions WHERE id = ?", (sid,)).fetchone()
    return _row_to_dict(row)


def update_session(sid: str, **kwargs) -> Optional[Dict]:
    """
    Update arbitrary fields on a session.

    Accepts any column name as a keyword argument.  JSON-serialisable values
    for *_json columns should be passed as Python objects (list/dict) — they
    will be serialised automatically.
    """
    if not kwargs:
        return get_session(sid)

    # Auto-serialise JSON fields
    json_fields = {"entries_json", "groups_json", "combos_json", "dst_programs_json"}
    for key in json_fields:
        if key in kwargs and not isinstance(kwargs[key], str):
            kwargs[key] = json.dumps(kwargs[key])

    # Always bump updated_at
    kwargs["updated_at"] = _now()

    set_clause = ", ".join(f"{k} = ?" for k in kwargs)
    values = list(kwargs.values())
    values.append(sid)

    db = get_db()
    db.execute(f"UPDATE sessions SET {set_clause} WHERE id = ?", values)
    db.commit()
    return get_session(sid)


def delete_session(sid: str) -> bool:
    """Delete a session row. Returns True if a row was deleted."""
    db = get_db()
    cur = db.execute("DELETE FROM sessions WHERE id = ?", (sid,))
    db.commit()
    return cur.rowcount > 0


def list_sessions(user_email: Optional[str] = None) -> List[Dict]:
    """
    List sessions as lightweight summary dicts.

    Returns: id, name, created_at, updated_at, has_excel, entries_count,
             combo_count, dst_count, exported
    """
    db = get_db()
    if user_email:
        rows = db.execute(
            "SELECT * FROM sessions WHERE user_email = ? ORDER BY updated_at DESC",
            (user_email,),
        ).fetchall()
    else:
        rows = db.execute(
            "SELECT * FROM sessions ORDER BY updated_at DESC"
        ).fetchall()

    result = []
    for row in rows:
        d = dict(row)
        result.append({
            "id": d["id"],
            "name": d["name"],
            "created_at": d["created_at"],
            "updated_at": d["updated_at"],
            "has_excel": d.get("excel_filename") is not None,
            "entries_count": _json_len(d.get("entries_json")),
            "combo_count": _json_len(d.get("combos_json")),
            "dst_count": _json_len(d.get("dst_programs_json")),
            "exported": bool(d.get("exported", 0)),
        })
    return result


# ---------------------------------------------------------------------------
# Session directory helpers
# ---------------------------------------------------------------------------

def get_session_dir(sid: str) -> str:
    """
    Return the base directory for a session's files, creating it and its
    subdirectories (dst/, output/) if they don't already exist.
    """
    session_dir = os.path.join(SESSIONS_DIR, sid)
    os.makedirs(os.path.join(session_dir, "dst"), exist_ok=True)
    os.makedirs(os.path.join(session_dir, "output"), exist_ok=True)
    return session_dir


def get_dst_dir(sid: str) -> str:
    """Convenience: return {session_dir}/dst/."""
    return os.path.join(get_session_dir(sid), "dst")


def get_output_dir(sid: str) -> str:
    """Convenience: return {session_dir}/output/."""
    return os.path.join(get_session_dir(sid), "output")
