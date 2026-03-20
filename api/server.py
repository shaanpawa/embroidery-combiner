"""
FastAPI backend for Micro Automation — Combo Builder.
Handles Excel parsing, DST file uploads, and combo export.
Sessions are persisted in SQLite via database.py.
"""

import json
import os
import shutil
import uuid
import zipfile
from datetime import datetime, timezone
from io import BytesIO

from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.core.excel_parser import (
    NameEntry, parse_excel, detect_columns, group_entries, generate_all_combos,
)
from app.core.pipeline import export_all
from api.database import (
    get_db, create_session, get_session, update_session,
    delete_session as db_delete_session, list_sessions,
    get_session_dir, get_dst_dir, get_output_dir,
)
from api.auth import get_current_user

app = FastAPI(title="Micro Automation API")


@app.get("/api/health")
async def health():
    """Health check — also used to wake up Render free tier before user needs it."""
    return {"status": "ok"}


# CORS: allow localhost for dev + production frontend URL from env
_cors_origins = ["http://localhost:3000", "http://localhost:5123"]
_frontend_url = os.environ.get("FRONTEND_URL")
if _frontend_url:
    _cors_origins.append(_frontend_url)

app.add_middleware(
    CORSMiddleware,
    allow_origins=_cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _require_session(session_id: str) -> dict:
    session = get_session(session_id)
    if session is None:
        raise HTTPException(404, "Session not found")
    return session


def _ensure_session(session_id: str = None, user_email: str = "local@dev") -> tuple:
    if session_id:
        session = get_session(session_id)
        if session:
            return session_id, session
    sid = session_id or str(uuid.uuid4())[:8]
    session = create_session(sid, sid, user_email)
    return sid, session


def _entries_to_json(entries: list) -> list:
    return [
        {
            "program": e.program,
            "name_line1": e.name_line1,
            "name_line2": e.name_line2,
            "quantity": e.quantity,
            "com_no": e.com_no,
            "machine_program": e.machine_program,
        }
        for e in entries
    ]


def _entries_from_json(data: list) -> list:
    return [
        NameEntry(
            program=d["program"],
            name_line1=d["name_line1"],
            name_line2=d["name_line2"],
            quantity=d["quantity"],
            com_no=d["com_no"],
            machine_program=d["machine_program"],
        )
        for d in data
    ]


def _combos_to_json(combos: list) -> list:
    return [
        {
            "filename": c.filename,
            "machine_program": c.machine_program,
            "com_no": c.com_no,
            "part_number": c.part_number,
            "total_parts": c.total_parts,
            "slot_count": len(c.slots),
            "left_count": len(c.left_column),
            "right_count": len(c.right_column),
            "slots": [
                {
                    "program": s.program,
                    "name_line1": s.name_line1,
                    "name_line2": s.name_line2,
                    "quantity": s.quantity,
                    "com_no": s.com_no,
                    "machine_program": s.machine_program,
                }
                for s in c.slots
            ],
        }
        for c in combos
    ]


def _build_groups_response(entries, combos, groups) -> list:
    groups_data = []
    for g in groups:
        g_combos = [c for c in combos if c.machine_program == g.machine_program and c.com_no == g.com_no]
        groups_data.append({
            "machine_program": g.machine_program,
            "com_no": g.com_no,
            "entry_count": len(g.entries),
            "total_slots": g.total_slots,
            "combos": [
                {
                    "filename": c.filename,
                    "part_number": c.part_number,
                    "total_parts": c.total_parts,
                    "slot_count": len(c.slots),
                    "left_count": len(c.left_column),
                    "right_count": len(c.right_column),
                    "slots": [
                        {"program": s.program, "name_line1": s.name_line1,
                         "name_line2": s.name_line2, "quantity": s.quantity}
                        for s in c.slots
                    ],
                }
                for c in g_combos
            ],
        })
    return groups_data


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.post("/api/detect-columns")
async def detect_columns_endpoint(
    file: UploadFile = File(...),
    session_id: str = Form(None),
    user: dict = Depends(get_current_user),
):
    """Upload Excel and auto-detect column mapping. Returns mapping + preview for confirmation."""
    sid, session = _ensure_session(session_id, user["email"])

    session_dir = get_session_dir(sid)
    excel_path = os.path.join(session_dir, file.filename or "order.xlsx")
    with open(excel_path, "wb") as f:
        content = await file.read()
        f.write(content)

    update_session(sid, excel_filename=file.filename)

    detection = detect_columns(excel_path)

    return {
        "session_id": sid,
        "excel_filename": file.filename,
        "headers": detection.headers,
        "preview_rows": detection.preview_rows,
        "detected_mapping": detection.detected_mapping,
        "confidence": detection.confidence,
    }


@app.post("/api/parse-excel")
async def parse_excel_endpoint(
    file: UploadFile = File(None),
    session_id: str = Form(None),
    column_map: str = Form(None),
    user: dict = Depends(get_current_user),
):
    sid, session = _ensure_session(session_id, user["email"])

    # Parse column_map if provided
    cmap = None
    if column_map:
        try:
            cmap = json.loads(column_map)
            # Ensure all values are ints
            cmap = {k: int(v) for k, v in cmap.items()}
        except (json.JSONDecodeError, ValueError):
            raise HTTPException(400, "Invalid column_map JSON")

    # If a new file is uploaded, save it. Otherwise use existing Excel in session.
    session_dir = get_session_dir(sid)
    if file:
        excel_path = os.path.join(session_dir, file.filename or "order.xlsx")
        with open(excel_path, "wb") as f:
            content = await file.read()
            f.write(content)
        update_session(sid, excel_filename=file.filename)
    else:
        # Find existing Excel file in session dir
        excel_files = [f for f in os.listdir(session_dir) if f.endswith(('.xlsx', '.xls'))]
        if not excel_files:
            raise HTTPException(400, "No Excel file found. Upload a file or use detect-columns first.")
        excel_path = os.path.join(session_dir, excel_files[0])

    result = parse_excel(excel_path, column_map=cmap)
    if not result.entries:
        update_session(sid, entries_json=[], groups_json=[], combos_json=[])
        return {"session_id": sid, "entries_count": 0, "groups": [], "combos": [], "warnings": result.warnings}

    groups = group_entries(result.entries)
    combos = generate_all_combos(result.entries)

    entries_data = _entries_to_json(result.entries)
    groups_data = _build_groups_response(result.entries, combos, groups)
    combos_data = _combos_to_json(combos)

    update_session(
        sid,
        entries_json=entries_data,
        groups_json=groups_data,
        combos_json=combos_data,
    )

    return {
        "session_id": sid,
        "entries_count": len(result.entries),
        "total_slots": sum(e.quantity for e in result.entries),
        "groups": groups_data,
        "combo_count": len(combos),
        "warnings": result.warnings,
        "entries_preview": entries_data,
    }


@app.post("/api/upload-dst")
async def upload_dst_endpoint(
    session_id: str = Form(...),
    files: list[UploadFile] = File(None),
    zip_file: UploadFile = File(None),
    user: dict = Depends(get_current_user),
):
    _require_session(session_id)
    dst_dir = get_dst_dir(session_id)
    found_programs = []

    if zip_file and zip_file.filename and zip_file.filename.lower().endswith(".zip"):
        content = await zip_file.read()
        with zipfile.ZipFile(BytesIO(content)) as zf:
            for name in zf.namelist():
                basename = os.path.basename(name)
                if basename.lower().endswith(".dst") and not basename.startswith("."):
                    target = os.path.join(dst_dir, basename)
                    with open(target, "wb") as out:
                        out.write(zf.read(name))

    if files:
        for f in files:
            if f.filename and f.filename.lower().endswith(".dst"):
                target = os.path.join(dst_dir, f.filename)
                with open(target, "wb") as out:
                    content = await f.read()
                    out.write(content)

    for fname in os.listdir(dst_dir):
        if fname.lower().endswith(".dst"):
            stem = os.path.splitext(fname)[0]
            try:
                found_programs.append(int(stem))
            except ValueError:
                pass

    found_programs.sort()
    session = update_session(session_id, dst_programs_json=found_programs)

    entries_data = session.get("entries_json") or []
    needed = set(e["program"] for e in entries_data)
    found_set = set(found_programs)
    missing = sorted(needed - found_set)

    return {
        "session_id": session_id,
        "uploaded_count": len(found_programs),
        "found_programs": found_programs,
        "needed_count": len(needed),
        "missing_programs": missing,
        "all_matched": len(missing) == 0,
    }


@app.post("/api/export")
async def export_endpoint(
    session_id: str = Form(...),
    selected_filenames: str = Form(""),
    gap_mm: float = Form(3.0),
    column_gap_mm: float = Form(5.0),
    user: dict = Depends(get_current_user),
):
    session = _require_session(session_id)
    entries_data = session.get("entries_json") or []
    if not entries_data:
        raise HTTPException(400, "No combos parsed yet")

    entries = _entries_from_json(entries_data)
    combos = generate_all_combos(entries)

    selected_set = set(f.strip() for f in selected_filenames.split(",") if f.strip())
    if selected_set:
        combos_to_export = [c for c in combos if c.filename in selected_set]
    else:
        combos_to_export = combos

    if not combos_to_export:
        raise HTTPException(400, "No combos selected")

    import tempfile

    output_dir = get_output_dir(session_id)
    dst_dir = get_dst_dir(session_id)

    # Export to temp dir first — only replace output_dir on success
    tmp_dir = tempfile.mkdtemp(prefix="combo_export_")
    try:
        results = export_all(
            combos_to_export, dst_dir, tmp_dir,
            gap_mm=gap_mm, column_gap_mm=column_gap_mm, overwrite=True,
        )

        success = [r for r in results if r.success]
        failed = [r for r in results if not r.success]

        if not success:
            raise HTTPException(500, f"All exports failed. First error: {failed[0].error if failed else 'unknown'}")

        # Success — move files from temp to output dir
        for f in os.listdir(output_dir):
            os.remove(os.path.join(output_dir, f))
        for r in success:
            dest = os.path.join(output_dir, os.path.basename(r.output_path))
            shutil.move(r.output_path, dest)
            r.output_path = dest
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    update_session(
        session_id,
        exported=1,
        exported_at=datetime.now(timezone.utc).isoformat(),
        gap_mm=gap_mm,
        column_gap_mm=column_gap_mm,
    )

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in success:
            zf.write(r.output_path, os.path.basename(r.output_path))
    zip_buffer.seek(0)

    session_name = session.get("name", session_id)
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition": f"attachment; filename=combos_{session_name}.zip",
            "X-Export-Success": str(len(success)),
            "X-Export-Failed": str(len(failed)),
        },
    )


# ---------------------------------------------------------------------------
# Session management
# ---------------------------------------------------------------------------

@app.get("/api/session/{session_id}/status")
async def session_status(session_id: str, user: dict = Depends(get_current_user)):
    session = _require_session(session_id)
    entries = session.get("entries_json") or []
    combos = session.get("combos_json") or []
    dst = session.get("dst_programs_json") or []
    return {
        "session_id": session_id,
        "has_excel": len(entries) > 0,
        "entries_count": len(entries),
        "combo_count": len(combos),
        "dst_count": len(dst),
        "exported": session.get("exported", False),
    }


@app.get("/api/session/{session_id}/full")
async def session_full(session_id: str, user: dict = Depends(get_current_user)):
    """Load complete session state for switching between sessions."""
    session = _require_session(session_id)
    entries = session.get("entries_json") or []
    groups = session.get("groups_json") or []
    dst_programs = session.get("dst_programs_json") or []

    needed = set(e["program"] for e in entries)
    found_set = set(dst_programs)
    missing = sorted(needed - found_set)

    return {
        "session_id": session_id,
        "name": session["name"],
        "entries_count": len(entries),
        "total_slots": sum(e.get("quantity", 1) for e in entries),
        "groups": groups,
        "combo_count": len(session.get("combos_json") or []),
        "entries_preview": entries,
        "dst_count": len(dst_programs),
        "dst_all_matched": len(missing) == 0,
        "missing_programs": missing,
        "exported": session.get("exported", False),
        "exported_at": session.get("exported_at"),
        "excel_filename": session.get("excel_filename"),
        "gap_mm": session.get("gap_mm", 3.0),
        "column_gap_mm": session.get("column_gap_mm", 5.0),
        "warnings": [],
    }


@app.get("/api/sessions")
async def list_sessions_endpoint(user: dict = Depends(get_current_user)):
    sessions = list_sessions(user["email"])
    for s in sessions:
        s["session_id"] = s.pop("id")
    return sessions


@app.post("/api/session/name")
async def set_session_name(
    session_id: str = Form(...), name: str = Form(...),
    user: dict = Depends(get_current_user),
):
    _require_session(session_id)
    update_session(session_id, name=name)
    return {"ok": True}


@app.delete("/api/session/{session_id}")
async def delete_session_endpoint(session_id: str, user: dict = Depends(get_current_user)):
    _require_session(session_id)
    session_dir = get_session_dir(session_id)
    if os.path.isdir(session_dir):
        shutil.rmtree(session_dir, ignore_errors=True)
    db_delete_session(session_id)
    return {"ok": True}


@app.delete("/api/session/{session_id}/excel")
async def remove_excel(session_id: str, user: dict = Depends(get_current_user)):
    """Remove Excel data, resetting parsed entries/combos/groups."""
    _require_session(session_id)
    update_session(
        session_id,
        excel_filename=None,
        entries_json=[],
        groups_json=[],
        combos_json=[],
        exported=0,
        exported_at=None,
    )
    output_dir = get_output_dir(session_id)
    if os.path.isdir(output_dir):
        for f in os.listdir(output_dir):
            os.remove(os.path.join(output_dir, f))
    return {"ok": True}


# ---------------------------------------------------------------------------
# Dev endpoint
# ---------------------------------------------------------------------------

@app.get("/api/dev/load-sample")
async def dev_load_sample():
    """Dev endpoint: auto-load real Excel + DST files for testing."""
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    test_data = os.path.join(project_root, "test_data")
    excel_path = os.path.join(test_data, "nameorder_04032026-2 add column (1).xlsx")
    if not os.path.isfile(excel_path):
        excel_path = os.path.expanduser("~/Downloads/nameorder_04032026-2 add column (1).xlsx")
    if not os.path.isfile(excel_path):
        excel_path = os.path.expanduser("~/Downloads/nameorder_04032026-2 add column.xlsx")
    dst_zip = os.path.join(test_data, "programs_Micro.zip")
    if not os.path.isfile(dst_zip):
        dst_zip = os.path.expanduser("~/Downloads/programs_Micro.zip")

    if not os.path.isfile(excel_path):
        raise HTTPException(404, f"Excel not found: {excel_path}")

    sid, session = _ensure_session("dev")

    result = parse_excel(excel_path)
    groups = group_entries(result.entries)
    combos = generate_all_combos(result.entries)

    entries_data = _entries_to_json(result.entries)
    groups_data = _build_groups_response(result.entries, combos, groups)
    combos_data = _combos_to_json(combos)

    update_session(
        sid,
        name="Dev Sample",
        excel_filename=os.path.basename(excel_path),
        entries_json=entries_data,
        groups_json=groups_data,
        combos_json=combos_data,
    )

    dst_dir = get_dst_dir(sid)
    if os.path.isfile(dst_zip):
        import zipfile as zf_mod
        with zf_mod.ZipFile(dst_zip) as z:
            for name in z.namelist():
                basename = os.path.basename(name)
                if basename.lower().endswith(".dst") and not basename.startswith("."):
                    with open(os.path.join(dst_dir, basename), "wb") as out:
                        out.write(z.read(name))

    dst_folder = os.path.join(project_root, "test_real_dst")
    if os.path.isdir(dst_folder):
        for fname in os.listdir(dst_folder):
            if fname.lower().endswith(".dst"):
                src = os.path.join(dst_folder, fname)
                tgt = os.path.join(dst_dir, fname)
                if not os.path.exists(tgt):
                    shutil.copy2(src, tgt)

    found = []
    for fname in os.listdir(dst_dir):
        if fname.lower().endswith(".dst"):
            try:
                found.append(int(os.path.splitext(fname)[0]))
            except ValueError:
                pass
    found.sort()
    update_session(sid, dst_programs_json=found)

    return {
        "session_id": sid,
        "entries_count": len(result.entries),
        "total_slots": sum(e.quantity for e in result.entries),
        "groups": groups_data,
        "combo_count": len(combos),
        "warnings": result.warnings,
        "entries_preview": entries_data,
        "dst_count": len(found),
        "dst_all_matched": True,
    }
