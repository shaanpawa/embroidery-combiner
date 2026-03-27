"""
SQLite persistence layer for Micro Automation — Embroidery Stacker.
Replaces the in-memory sessions dict in server.py.

DB file: {project_root}/data/micro.db
Session files: {project_root}/data/sessions/{sid}/dst/ and output/
"""

import json
import os
import sqlite3
import threading
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.environ.get("MICRO_DATA_DIR", os.path.join(PROJECT_ROOT, "data"))
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
    conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=10.0)
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
            exported_at      TEXT,
            assign_result_json TEXT
        )
    """)
    # Migrate: add columns if they don't exist (for existing DBs)
    for col, col_type in [
        ("assign_result_json", "TEXT"),
        ("optimize_heads", "INTEGER DEFAULT 0"),
    ]:
        try:
            conn.execute(f"ALTER TABLE sessions ADD COLUMN {col} {col_type}")
            conn.commit()
        except sqlite3.OperationalError:
            pass  # column already exists
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ma_reference (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            size_normalized  TEXT NOT NULL UNIQUE,
            size_display     TEXT NOT NULL,
            ma_number        TEXT NOT NULL,
            created_at       TEXT NOT NULL DEFAULT (datetime('now'))
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS com_reference (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            ma_number        TEXT NOT NULL,
            com_number       INTEGER NOT NULL,
            fabric_colour    TEXT NOT NULL,
            embroidery_colour TEXT NOT NULL,
            frame_colour     TEXT NOT NULL,
            created_at       TEXT NOT NULL DEFAULT (datetime('now')),
            UNIQUE(ma_number, fabric_colour, embroidery_colour, frame_colour)
        )
    """)
    conn.commit()
    return conn


# ---------------------------------------------------------------------------
# MA Reference CRUD
# ---------------------------------------------------------------------------

def get_ma_reference() -> List[Dict]:
    """Return all MA reference mappings as a list of dicts."""
    db = get_db()
    rows = db.execute(
        "SELECT size_normalized, size_display, ma_number FROM ma_reference ORDER BY size_display"
    ).fetchall()
    return [dict(r) for r in rows]


def get_ma_lookup() -> Dict[str, str]:
    """Return a dict of normalized_size → ma_number for use in auto-assign."""
    db = get_db()
    rows = db.execute(
        "SELECT size_normalized, ma_number FROM ma_reference"
    ).fetchall()
    return {r["size_normalized"]: r["ma_number"] for r in rows}


def upsert_ma_reference(mappings: List[Dict]) -> int:
    """Insert or replace MA reference mappings.

    Each mapping must have: size_normalized, size_display, ma_number.
    Returns the number of rows upserted.
    """
    db = get_db()
    now = _now()
    with _db_lock:
        for m in mappings:
            db.execute(
                """
                INSERT INTO ma_reference (size_normalized, size_display, ma_number, created_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(size_normalized) DO UPDATE SET
                    size_display = excluded.size_display,
                    ma_number = excluded.ma_number,
                    created_at = excluded.created_at
                """,
                (m["size_normalized"], m["size_display"], m["ma_number"], now),
            )
        db.commit()
    return len(mappings)


def clear_ma_reference() -> int:
    """Delete all MA reference mappings. Returns number of rows deleted."""
    db = get_db()
    with _db_lock:
        cur = db.execute("DELETE FROM ma_reference")
        db.commit()
    return cur.rowcount


# ---------------------------------------------------------------------------
# COM Reference CRUD
# ---------------------------------------------------------------------------

def get_com_reference() -> List[Dict]:
    """Return all COM reference mappings as a list of dicts."""
    db = get_db()
    rows = db.execute(
        "SELECT ma_number, com_number, fabric_colour, embroidery_colour, frame_colour "
        "FROM com_reference ORDER BY ma_number, com_number"
    ).fetchall()
    return [dict(r) for r in rows]


def get_com_lookup() -> Dict[str, Dict[tuple, int]]:
    """Return nested lookup: ma_number → {(fabric, embroidery, frame): com_number}.

    Color keys are title-cased to match auto-assign normalization.
    """
    db = get_db()
    rows = db.execute(
        "SELECT ma_number, com_number, fabric_colour, embroidery_colour, frame_colour "
        "FROM com_reference"
    ).fetchall()
    lookup: Dict[str, Dict[tuple, int]] = {}
    for r in rows:
        ma = r["ma_number"]
        key = (r["fabric_colour"], r["frame_colour"], r["embroidery_colour"])
        if ma not in lookup:
            lookup[ma] = {}
        lookup[ma][key] = r["com_number"]
    return lookup


def get_max_com_per_ma() -> Dict[str, int]:
    """Return the highest COM number per MA, so new COMs can continue sequencing."""
    db = get_db()
    rows = db.execute(
        "SELECT ma_number, MAX(com_number) as max_com FROM com_reference GROUP BY ma_number"
    ).fetchall()
    return {r["ma_number"]: r["max_com"] for r in rows}


def upsert_com_reference(mappings: List[Dict]) -> int:
    """Insert or replace COM reference mappings.

    Each mapping must have: ma_number, com_number, fabric_colour, embroidery_colour, frame_colour.
    """
    db = get_db()
    now = _now()
    with _db_lock:
        for m in mappings:
            db.execute(
                """
                INSERT INTO com_reference (ma_number, com_number, fabric_colour, embroidery_colour, frame_colour, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(ma_number, fabric_colour, embroidery_colour, frame_colour) DO UPDATE SET
                    com_number = excluded.com_number,
                    created_at = excluded.created_at
                """,
                (m["ma_number"], m["com_number"], m["fabric_colour"], m["embroidery_colour"], m["frame_colour"], now),
            )
        db.commit()
    return len(mappings)


def clear_com_reference() -> int:
    """Delete all COM reference mappings. Returns number of rows deleted."""
    db = get_db()
    with _db_lock:
        cur = db.execute("DELETE FROM com_reference")
        db.commit()
    return cur.rowcount


def add_single_ma(size_normalized: str, size_display: str, ma_number: str) -> Dict:
    """Add or update a single MA reference entry. Returns the saved entry."""
    db = get_db()
    now = _now()
    with _db_lock:
        db.execute(
            """
            INSERT INTO ma_reference (size_normalized, size_display, ma_number, created_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(size_normalized) DO UPDATE SET
                size_display = excluded.size_display,
                ma_number = excluded.ma_number,
                created_at = excluded.created_at
            """,
            (size_normalized, size_display, ma_number, now),
        )
        db.commit()
    row = db.execute(
        "SELECT id, size_normalized, size_display, ma_number FROM ma_reference WHERE size_normalized = ?",
        (size_normalized,),
    ).fetchone()
    return dict(row)


def update_ma_reference_entry(entry_id: int, **kwargs) -> Optional[Dict]:
    """Update a single MA reference entry by ID. Returns updated entry or None."""
    allowed = {"size_normalized", "size_display", "ma_number"}
    filtered = {k: v for k, v in kwargs.items() if k in allowed}
    if not filtered:
        return None
    db = get_db()
    set_clause = ", ".join(f"{k} = ?" for k in filtered)
    values = list(filtered.values()) + [entry_id]
    with _db_lock:
        db.execute(f"UPDATE ma_reference SET {set_clause} WHERE id = ?", values)
        db.commit()
    row = db.execute("SELECT id, size_normalized, size_display, ma_number FROM ma_reference WHERE id = ?", (entry_id,)).fetchone()
    return dict(row) if row else None


def delete_ma_reference_entry(entry_id: int) -> bool:
    """Delete a single MA reference entry by ID."""
    db = get_db()
    with _db_lock:
        cur = db.execute("DELETE FROM ma_reference WHERE id = ?", (entry_id,))
        db.commit()
    return cur.rowcount > 0


def add_single_com(ma_number: str, com_number: int, fabric_colour: str, embroidery_colour: str, frame_colour: str) -> Dict:
    """Add or update a single COM reference entry. Returns the saved entry."""
    db = get_db()
    now = _now()
    with _db_lock:
        db.execute(
            """
            INSERT INTO com_reference (ma_number, com_number, fabric_colour, embroidery_colour, frame_colour, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(ma_number, fabric_colour, embroidery_colour, frame_colour) DO UPDATE SET
                com_number = excluded.com_number,
                created_at = excluded.created_at
            """,
            (ma_number, com_number, fabric_colour, embroidery_colour, frame_colour, now),
        )
        db.commit()
    row = db.execute(
        "SELECT id, ma_number, com_number, fabric_colour, embroidery_colour, frame_colour "
        "FROM com_reference WHERE ma_number = ? AND fabric_colour = ? AND embroidery_colour = ? AND frame_colour = ?",
        (ma_number, fabric_colour, embroidery_colour, frame_colour),
    ).fetchone()
    return dict(row)


def update_com_reference_entry(entry_id: int, **kwargs) -> Optional[Dict]:
    """Update a single COM reference entry by ID. Returns updated entry or None."""
    allowed = {"ma_number", "com_number", "fabric_colour", "embroidery_colour", "frame_colour"}
    filtered = {k: v for k, v in kwargs.items() if k in allowed}
    if not filtered:
        return None
    db = get_db()
    set_clause = ", ".join(f"{k} = ?" for k in filtered)
    values = list(filtered.values()) + [entry_id]
    with _db_lock:
        db.execute(f"UPDATE com_reference SET {set_clause} WHERE id = ?", values)
        db.commit()
    row = db.execute(
        "SELECT id, ma_number, com_number, fabric_colour, embroidery_colour, frame_colour FROM com_reference WHERE id = ?",
        (entry_id,),
    ).fetchone()
    return dict(row) if row else None


def delete_com_reference_entry(entry_id: int) -> bool:
    """Delete a single COM reference entry by ID."""
    db = get_db()
    with _db_lock:
        cur = db.execute("DELETE FROM com_reference WHERE id = ?", (entry_id,))
        db.commit()
    return cur.rowcount > 0


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
    # Deserialise assign_result_json as dict (not list)
    raw_assign = d.get("assign_result_json")
    d["assign_result_json"] = json.loads(raw_assign) if raw_assign else None
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
    with _db_lock:
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


_ALLOWED_COLUMNS = {
    "user_email", "name", "updated_at",
    "excel_filename", "entries_json", "groups_json", "combos_json",
    "dst_programs_json", "gap_mm", "column_gap_mm",
    "exported", "exported_at", "assign_result_json",
    "optimize_heads",
}


def update_session(sid: str, **kwargs) -> Optional[Dict]:
    """
    Update arbitrary fields on a session.

    Accepts any column name as a keyword argument.  JSON-serialisable values
    for *_json columns should be passed as Python objects (list/dict) — they
    will be serialised automatically.
    """
    if not kwargs:
        return get_session(sid)

    # Validate column names against whitelist to prevent SQL injection
    invalid_keys = set(kwargs.keys()) - _ALLOWED_COLUMNS
    if invalid_keys:
        raise ValueError(f"Invalid column names: {invalid_keys}")

    # Auto-serialise JSON fields
    json_fields = {"entries_json", "groups_json", "combos_json", "dst_programs_json", "assign_result_json"}
    for key in json_fields:
        if key in kwargs and not isinstance(kwargs[key], str):
            kwargs[key] = json.dumps(kwargs[key])

    # Always bump updated_at
    kwargs["updated_at"] = _now()

    set_clause = ", ".join(f"{k} = ?" for k in kwargs)
    values = list(kwargs.values())
    values.append(sid)

    db = get_db()
    with _db_lock:
        db.execute(f"UPDATE sessions SET {set_clause} WHERE id = ?", values)
        db.commit()
    return get_session(sid)


def delete_session(sid: str) -> bool:
    """Delete a session row. Returns True if a row was deleted."""
    db = get_db()
    with _db_lock:
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
        # Calculate expiry time (updated_at + 24 hours)
        expires_at = None
        try:
            updated = datetime.fromisoformat(d["updated_at"])
            expires_at = (updated + timedelta(hours=24)).isoformat()
        except (ValueError, TypeError):
            pass
        result.append({
            "id": d["id"],
            "name": d["name"],
            "created_at": d["created_at"],
            "updated_at": d["updated_at"],
            "expires_at": expires_at,
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


# ---------------------------------------------------------------------------
# Cleanup
# ---------------------------------------------------------------------------

def cleanup_old_sessions(max_age_hours: int = 24) -> int:
    """Delete sessions older than *max_age_hours* and their directories.

    Returns the number of sessions deleted.
    """
    import shutil

    cutoff = (datetime.now(timezone.utc) - timedelta(hours=max_age_hours)).isoformat()
    db = get_db()

    with _db_lock:
        rows = db.execute(
            "SELECT id FROM sessions WHERE updated_at < ?", (cutoff,)
        ).fetchall()
        deleted = 0
        for row in rows:
            sid = row["id"]
            session_dir = os.path.join(SESSIONS_DIR, sid)
            if os.path.isdir(session_dir):
                shutil.rmtree(session_dir, ignore_errors=True)
            db.execute("DELETE FROM sessions WHERE id = ?", (sid,))
            deleted += 1
        if deleted:
            db.commit()
    return deleted
