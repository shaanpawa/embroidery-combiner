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
  src/app/combo-builder/page.tsx — Combo Builder workflow UI
  src/app/login/page.tsx — Google OAuth login page
  src/auth.ts           — NextAuth config (Google OAuth, JWT)
  src/middleware.ts      — Route protection
  src/app/globals.css   — Design system (cream/green palette + glassmorphism)

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
- `NEXTAUTH_URL` — canonical URL
- `NEXTAUTH_SECRET` — shared JWT secret with backend
- `GOOGLE_CLIENT_ID` — Google OAuth client ID
- `GOOGLE_CLIENT_SECRET` — Google OAuth client secret
- `AUTH_TRUST_HOST` — set `true` for Vercel
- `NEXT_PUBLIC_API_URL` — Render backend URL

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

The Excel order sheet drives everything. Here's how columns map:

| Excel Column | Field | Purpose |
|--------------|-------|---------|
| A (Program) | `program` | The DST file number (e.g., "42" → matches `42.DST`) |
| F (row1) | `row1` | First name line |
| G (row2) | `row2` | Second name line (optional) |
| H (Quantity) | `quantity` | How many slots this name takes (qty=2 means 2 copies) |
| O (Com No) | `com_no` | Combo number — names with same (M, ComNo) go in same combo |
| P (Machine Program) | `machine_program` | Machine program code (e.g., "MA") — groups patch type |

**Grouping:** Names with the same `(machine_program, com_no)` pair are grouped together into combo files.

**Splitting:** If a group exceeds 20 slots, it splits into multiple combo files (e.g., Group1-1, Group1-2).

**NOTE:** If the Excel format differs (columns in different positions), the operator needs to correct the mapping. The current parser hardcodes column positions. A future enhancement should let the operator visually confirm/remap which columns map to which fields.

## Design System

- Background: warm cream (#f5f0eb) / dark (#0c0c0c)
- Accent: forest green (#2b5e49) / lighter green (#4ead8a)
- Typography: Geist Sans / Geist Mono
- Effects: Glassmorphism, liquid glass cards, backdrop-filter blur
- Brand: "Micro Automation by Ossia"
- Light/dark theme toggle

## Current Status (March 19, 2026)

### DONE:
- Core logic: excel parser, two-column combiner, pipeline — 81 tests passing
- Verified with real data: 300 names, 300 DST files → 31 combo files exported
- FastAPI backend: all endpoints working (parse, upload, export)
- SQLite persistence with session management
- Next.js frontend: launcher page + combo builder workflow
- Google OAuth authentication (NextAuth) with email whitelist
- Desktop app: functional (customtkinter)
- Fixed COLOR_CHANGE bug: no extra machine stops between names
- Fixed metadata: clean DST headers
- Deployed: Vercel (frontend) + Render (backend)

### PRIORITY ROADMAP:
1. **Responsive website** — mobile-friendly UI, works on tablets/phones for operators on factory floor
2. **Google auth with whitelist** — finalize Google OAuth setup, configure allowed emails for Micro operators
3. **Color/needle logic verification** — verify COLOR_CHANGE commands work correctly on actual machine (use MyEditor on Windows to inspect)
4. **Needle up handling** — investigate whether needle-up commands are needed between designs (test with MyEditor)
5. **Excel mapping clarity** — visualize how Excel columns map to combo logic (MA code → Combo code), let operators see and correct if Excel format differs
6. **Grouping validation view** — let operator see/override auto-grouped combos before export

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
- next-auth 5 (beta) — Google OAuth authentication
- Tailwind CSS 4 — styling
- Geist — typography
