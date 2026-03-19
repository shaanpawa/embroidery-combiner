# Micro Automation by Ossia — Combo Builder

Combines individual embroidery name programs (DST files) into production-ready combo files based on an Excel order sheet. Built for [Micro Embroidery Co.](https://www.microembroidery.com/) (Bangkok, Thailand).

## Live App

**https://embroidery-combiner.vercel.app**

## How It Works

1. **Upload Excel** — order sheet with program numbers, names, quantities, combo groups
2. **Upload DST files** — individual name program files (e.g., 1.DST through 300.DST)
3. **Review combos** — app auto-groups names into combo files based on Machine Program + Com No
4. **Export** — download a zip of combined DST files ready for the machine

## Run Locally

```bash
# Backend API
pip install -r requirements.txt
python -m uvicorn api.server:app --reload --port 8000

# Frontend (separate terminal)
cd web && npm install && npm run dev    # localhost:3000
```

Set `AUTH_DISABLED=true` in `.env` to skip authentication during development.

## Tests

```bash
pytest tests/ -v    # 81 tests
```

## Architecture

- **Backend:** FastAPI + SQLite (api/)
- **Frontend:** Next.js + React + Tailwind (web/)
- **Core logic:** Python (app/core/) — excel parsing, DST combining, pipeline
- **Desktop app:** customtkinter (app/ui/) — functional but web is primary

See [PROJECT_CONTEXT.md](PROJECT_CONTEXT.md) for full technical details.
