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

web/                    — Next.js frontend (port 3000)
  src/app/page.tsx      — Product launcher ("Micro Automation by Ossia")
  src/app/combo-builder/page.tsx — Combo Builder workflow
  src/app/globals.css   — Design system (ScrapYard palette + liquid glass)

app/ui/                 — Desktop UI (customtkinter) — functional but web is primary path
  combo_app.py          — Excel-driven workflow
  theme.py              — Dark/light design system

tests/                  — 81 automated tests (pytest)
_archive/               — Legacy files (pre-refactor standalone modules, old builds)
```

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

## Design System

Based on ScrapYard project palette:
- Background: warm cream (#f5f0eb)
- Accent: forest green (#2b5e49)
- Typography: Geist Sans / Geist Mono
- Effects: Glassmorphism, liquid glass cards, backdrop-filter blur
- Brand: "Micro Automation by Ossia"

## Current Status (March 19, 2026)

### DONE:
- Core logic: excel parser, two-column combiner, pipeline — 81 tests passing
- Verified with real data: 300 names, 300 DST files → 31 combo files exported
- FastAPI backend: all endpoints working (parse, upload, export)
- Next.js frontend: launcher page + combo builder page rendering
- Desktop app: functional (customtkinter)
- Fixed COLOR_CHANGE bug: no extra machine stops between names
- Fixed metadata: clean DST headers
- Full batch export tested: 31 combo files generated, ready for machine testing

### IN PROGRESS:
- Machine testing: combo files need to be tested on actual embroidery machine
- Web app end-to-end testing (file upload flow in browser needs manual testing)

### NOT TESTED:
- NGS→DST conversion via pywinauto (needs Windows)
- Desktop app on Windows

## How to Run

```bash
# Tests
cd /Users/shaan_pawa/Micro/embroidery-combiner
source .venv/bin/activate
pytest tests/ -v                # 81 tests

# Web app (both servers needed)
python -m uvicorn api.server:app --reload --port 8000    # API
cd web && npm run dev                                     # Frontend at localhost:3000

# Desktop app
python main.py

# Dev shortcut: load real test data via API
curl http://localhost:8000/api/dev/load-sample
```

## Real Test Data
- Excel: `test_data/nameorder_04032026-2 add column (1).xlsx` (300 names, 15 groups, 31 combos)
- DST zip: `test_data/programs_Micro.zip` (300 DST files, 1.DST-300.DST)
- NGS reference: `test_data/170-183คอม4ปักหัว.ngs` (manually-created correct combo for comparison)
- Extracted DSTs: `test_real_dst/`
- Dev endpoint loads from `test_data/` — see `api/server.py` `/api/dev/load-sample`
