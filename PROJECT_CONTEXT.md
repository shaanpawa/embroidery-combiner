# Micro Automation by Ossia — Project Context

## What This Is

**Micro Automation** is a product suite by Ossia that automates workflows for [Micro Embroidery Co.](https://www.microembroidery.com/) (Bangkok, Thailand — manufacturer of embroidery badges/patches).

**First product: Combo Builder** — Combines individual embroidery name programs into production-ready combo files. An operator uploads an Excel order sheet + a folder of DST program files, and the app generates combined DST combo files ready for the machine.

## How Combo Builder Works

**Operator does:** Upload Excel order → Upload DST programs → Review combos → Export combined files

**App does:**
1. Parses Excel order (columns: Program number, Name, Title, Quantity, Com No, Machine Program)
2. Groups names by (Machine Program + Com No) — determines which names go in the same combo
3. Expands quantities (qty=2 = name takes 2 slots)
4. Splits groups into combo files of max 20 slots (10 left column + 10 right column)
5. Maps each slot to its DST file (by program number)
6. Combines DST files in two-column layout with configurable gaps
7. Exports as downloadable zip of combo DST files

## Architecture

```
app/core/               — Business logic (no UI dependencies, 81 tests)
  excel_parser.py       — Parse Excel, group by (M + ComNo), expand qty, split into combos
  combiner.py           — DST combining: single-col + two-col layout
  pipeline.py           — Maps combo slots → DST files, orchestrates export
  converter.py          — NGS→DST via GUI automation (Windows only, untested)
  file_discovery.py     — Folder scanning (legacy workflow, still available)
  validator.py          — File validation (DST/NGS)

app/config.py           — Settings persistence, constants (v2.0.0)
app/licensing.py        — Hardware-locked HMAC licensing

api/                    — FastAPI backend (port 8000)
  server.py             — Endpoints: parse-excel, upload-dst, export, dev/load-sample
  database.py           — SQLite persistence, session management
  auth.py               — NextAuth JWT validation, email whitelist

web/                    — Next.js frontend (port 3000)
  src/app/page.tsx      — Product launcher ("Micro Automation by Ossia")
  src/app/combo-builder/page.tsx — Combo Builder workflow UI (~980 lines)
  src/app/login/page.tsx — Password + Google login page
  src/app/i18n.tsx      — Thai/English language provider (~50 translation keys)
  src/auth.ts           — NextAuth config (Credentials + Google, JWT)
  src/middleware.ts      — Route protection
  src/lib/api.ts        — authFetch utility (Bearer JWT for cross-origin API calls)
  src/app/globals.css   — Design system (cream/blue palette + glassmorphism)

app/ui/                 — Desktop UI (customtkinter) — functional but web is primary
  combo_app.py          — Excel-driven workflow
  theme.py              — Dark/light design system

tests/                  — 81 automated tests (pytest)
```

## Deployment

| Component | Host | URL |
|-----------|------|-----|
| Frontend | Vercel | https://embroidery-combiner.vercel.app |
| Backend API | Render | (auto-deploy from main branch) |
| Desktop | Windows EXE | Built via PyInstaller |

Auto-deploy: push to `main` → Vercel rebuilds frontend, Render rebuilds backend.

### Environment Variables

**Backend (Render):**
- `AUTH_DISABLED` — set `true` to bypass auth (dev only)
- `ALLOWED_EMAILS` — comma-separated email whitelist for Google auth
- `NEXTAUTH_SECRET` — shared JWT secret with frontend
- `FRONTEND_URL` — Vercel URL for CORS

**Frontend (Vercel):**
- `ADMIN_PASSWORD` — password for Credentials login (currently: micro2026)
- `NEXTAUTH_SECRET` — shared JWT secret with backend
- `AUTH_TRUST_HOST` — set `true` for Vercel
- `NEXT_PUBLIC_API_URL` — Render backend URL
- `GOOGLE_CLIENT_ID` — Google OAuth client ID (not yet configured)
- `GOOGLE_CLIENT_SECRET` — Google OAuth client secret (not yet configured)

## Key Technical Facts

- **NGS is proprietary** — no library reads stitch data. pyembroidery does NOT support NGS. NGS stores color data; DST does not.
- **DST has zero color data** — viewers (TrueSizer, Wings) show rainbow palette colors which are purely cosmetic. Machine operator loads the correct threads (red + green).
- **DST combining works on any OS** — pyembroidery handles DST fully.
- **Critical END bug FIXED** — pyembroidery's DST writer stops at first END command. Fix: strip END between designs, add single END at end.
- **Critical COLOR_CHANGE bug FIXED** — pyembroidery's `add_pattern()` inserts extra COLOR_CHANGE commands between designs, causing machine stops. Fix: use `stitches.extend()` instead of `add_pattern()` to manually merge stitch data.
- **Combo file structure** — Each name has 1 COLOR_CHANGE (red→green, automatic needle switch). No STOP commands between names. TRIM between names for thread cutting. Single END at end of file.
- **Two-column layout** — 10 left + 10 right, 3mm vertical gap, 5mm horizontal gap (configurable)
- **Combo grouping** — (Machine Program, Com No). M defines patch size/type, Com No adds color grouping.
- **Case-insensitive file matching** — handles both `1.dst` and `1.DST`
- **Cannot byte-compare outputs** — visual verification by operator is the acceptance test.

## Excel Mapping Logic

The Excel order sheet drives everything. Column detection is **automatic** with operator confirmation:

1. **Auto-detect**: `detect_columns()` in `excel_parser.py` reads Excel headers and fuzzy-matches to 6 required fields using pattern matching (e.g., "program" → Program, "qty" → Quantity, "Com No" → Combo Number)
2. **Confirm**: Frontend shows a mapping UI with dropdowns + sample data so operator can verify/correct the auto-detected columns
3. **Parse**: `parse_excel()` accepts the confirmed `column_map` and processes the data

| Field | Purpose | Header Patterns |
|-------|---------|-----------------|
| `program` | DST file number (e.g., "42" → `42.DST`) | "program", "prog" |
| `name` | Name to be embroidered | "row 1", "name", "ชื่อ" |
| `title` | Second name line (optional) | "row 2", "title" |
| `quantity` | How many slots this name takes | "quantity", "qty", "จำนวน" |
| `com_no` | Combo number — same (M, ComNo) = same combo | "com no", "combo", "คอม" |
| `machine_program` | Machine program code (e.g., "MA50310") | single "M" column, "machine" |

**Grouping:** Names with the same `(machine_program, com_no)` pair are grouped together into combo files.

**Splitting:** If a group exceeds 20 slots, it splits into multiple combo files (e.g., Group1-1, Group1-2).

## Design System

- Background: warm cream (#f5f0eb) / dark (#0c0c0c)
- Accent: Micro blue (#26397A) / lighter blue (#6b8cdb)
- Typography: Geist Sans / Geist Mono
- Effects: Glassmorphism, liquid glass cards, backdrop-filter blur
- Brand: "Micro Automation by Ossia"
- Light/dark theme toggle
- Thai/English language toggle (persists to localStorage)

## Current Status (March 20, 2026)

### DONE:
- Core logic: excel parser, two-column combiner, pipeline — 81 tests passing
- Verified with real data: 300 names, 300 DST files → 31 combo files exported
- FastAPI backend: all endpoints working (parse, upload, export, detect-columns)
- SQLite persistence with session management (WAL mode, threading lock, 10s timeout)
- Next.js frontend: launcher page + combo builder workflow
- **Password authentication** (ADMIN_PASSWORD env var) + Google OAuth button (needs credentials)
- Cross-origin auth: NextAuth JWT → custom HS256 token → backend validation
- **Thai/English translation** (~50 keys, toggle in nav bar) — ~20 error strings still hardcoded English
- **Responsive UI** — mobile-friendly, 2x2 stat grid, touch-friendly items, slide-up preview overlay
- **Excel column auto-detection** — fuzzy header matching + confirmation UI with sample data
- **Stability fixes** — export writes to temp dir then swaps, 5-min token TTL, SQLite write lock
- Desktop app: functional (customtkinter)
- Fixed COLOR_CHANGE bug: no extra machine stops between names
- Fixed metadata: clean DST headers
- Deployed: Vercel (frontend) + Render (backend) — auto-deploy from `main`

### PRIORITY ROADMAP:
1. **Google OAuth setup** — create Google Cloud Console project, add GOOGLE_CLIENT_ID + GOOGLE_CLIENT_SECRET env vars, configure ALLOWED_EMAILS whitelist
2. **Hardening plan** — file size limits, fetch timeouts, finish Thai translations, button disable states, session cleanup, error recovery. Details in `.claude/plans/parsed-churning-twilight.md`
3. **Color/needle logic verification** — verify COLOR_CHANGE commands on actual machine (MyEditor on Windows)
4. **Needle up handling** — test whether needle-up commands needed between designs (MyEditor)

### NOT TESTED:
- NGS→DST conversion via pywinauto (needs Windows)
- Desktop app on Windows
- Color/needle behavior on actual machine

## How to Run

```bash
# Clone
git clone https://github.com/shaanpawa/embroidery-combiner.git
cd embroidery-combiner

# Backend
python -m venv .venv
source .venv/bin/activate          # Mac/Linux
# .venv\Scripts\activate           # Windows
pip install -r requirements.txt
python -m uvicorn api.server:app --reload --port 8000

# Frontend (separate terminal)
cd web
npm install
npm run dev                        # localhost:3000

# Tests
pytest tests/ -v                   # 81 tests

# Dev shortcut: load real test data via API
curl http://localhost:8000/api/dev/load-sample

# Desktop app (optional)
python main.py
```

### Local Dev Environment Variables

Create `web/.env.local`:
```
ADMIN_PASSWORD=micro2026
NEXTAUTH_SECRET=your-secret-here
NEXT_PUBLIC_API_URL=http://localhost:8000
AUTH_TRUST_HOST=true
# Optional (for Google OAuth):
# GOOGLE_CLIENT_ID=...
# GOOGLE_CLIENT_SECRET=...
```

Backend needs no env vars for local dev (auth disabled when no NEXTAUTH_SECRET set).

## Real Test Data
- Excel: `test_data/nameorder_04032026-2 add column (1).xlsx` (300 names, 15 groups, 31 combos)
- DST zip: `test_data/programs_Micro.zip` (300 DST files, 1.DST-300.DST)
- NGS reference: `test_data/170-183คอม4ปักหัว.ngs` (manually-created correct combo for comparison)
- Dev endpoint loads from `test_data/` — see `api/server.py` `/api/dev/load-sample`

## Dependencies

**Python (requirements.txt):**
- pyembroidery — DST reading/writing/combining
- openpyxl — Excel parsing
- fastapi + uvicorn — API server
- python-multipart — file uploads
- customtkinter — desktop UI (optional)
- pywinauto — NGS conversion automation (Windows only)

**JavaScript (web/package.json):**
- Next.js 16 + React 19 — frontend framework
- next-auth 5 (beta) — authentication (Credentials + Google OAuth)
- Tailwind CSS 4 — styling
- Geist — typography
