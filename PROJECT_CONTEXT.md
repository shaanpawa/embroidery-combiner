# Embroidery Combiner — Project Context

## What This Is

Tool for FM (Thai embroidery company). Automates combining sequential embroidery design files into one file, stacked vertically. Replaces a manual process that takes an operator 4-6 hours/week.

## How It Works

**Operator does:** Browse to folder → click Combine → get combined output file.

**App does behind the scenes:**
1. Scans folder for .ngs and .dst files
2. Sorts by number extracted from filename (216.ngs, 217.ngs, ...)
3. Validates each file (checks for corruption, empty files)
4. If NGS files found: converts each to DST via Wings software (pywinauto GUI automation, Windows only)
5. Combines all DST files vertically with configurable gap (default 3mm)
6. Saves combined output as `{first}-{last}.dst` (e.g., 216-225.dst) to same folder
7. Skips files that look like previous combined output (e.g., 216-225.dst won't be re-included)

## Architecture

```
app/core/           — Business logic (no UI dependencies, fully tested)
  combiner.py       — Reads DST files, stacks vertically, handles gaps and TRIM commands
  converter.py      — NGS→DST conversion via Wings XP or My Editor GUI automation (Windows only)
  file_discovery.py — Folder scanning, numeric sorting, sequence gap/duplicate detection
  validator.py      — File validation (DST via pyembroidery, NGS via OLE2 structure check)

app/config.py       — Settings persistence (JSON), app constants, gap presets
app/licensing.py    — Hardware-locked HMAC licensing (machine fingerprint from MAC+CPU+OS)

app/ui/             — Desktop UI (customtkinter) — built but NOT visually tested
  app.py            — Main window, pipeline orchestration, threading
  theme.py          — Colors, fonts, spacing
  components/       — Modular UI components (file_table, gap_controls, output_panel, etc.)

main.py             — Entry point (license check → launch app)
build.py            — Nuitka packaging for standalone .exe
web_demo.py         — Flask web demo (used for testing logic on Mac)
```

## Key Technical Facts

- **NGS is proprietary** — no library (including pyembroidery) can read NGS stitch data. Only Wings software (Windows) can decode it.
- **DST combining works on any OS** — pyembroidery handles DST fully.
- **95% of incoming files are NGS** — so in practice, this needs Windows.
- **Wings XP (paid) at FM, My Editor (free) for Shaan's testing** — converter auto-detects whichever is installed.
- **Cannot byte-compare outputs** — our DST output vs reference NGS files are different formats. Visual verification by operator is the acceptance test.

## Current Status (as of March 2026)

### WORKING (tested, verified):
- **Combiner logic** — 42 automated tests pass. Combines 2, 3, or 10 files. Write→read roundtrip preserves all designs.
- **Critical END bug FIXED** — pyembroidery adds END after each design; DST writer stops at first END. Fix: strip END between designs, add single END at end. (`app/core/combiner.py` lines 57-94)
- **File discovery** — scans folders, sorts by number, detects gaps/duplicates, skips previous output
- **Validation** — DST (pyembroidery parse), NGS (OLE2 structure check)
- **Web demo** — Flask app at web_demo.py, tested and working for DST combining on Mac

### TESTED ON MAC (March 2026):
- **Desktop UI launched and verified** — customtkinter app runs on Mac with Python 3.12 (Homebrew) + venv. Loads DST files, combines them, saves output. Progress indicator fixed (hides after completion).
- **Web demo tested** — Flask app at web_demo.py, confirmed working for DST combining on Mac.

### NOT TESTED (needs Windows):
- **NGS→DST conversion via pywinauto** — the converter code (`app/core/converter.py`) is written but the GUI automation has NEVER been tested against real Wings XP or My Editor. Menu names, button labels, dialog titles are all GUESSED based on standard Windows patterns. They must be verified on the actual software.
  - To verify: open Wings/My Editor on Windows, check if `"File->Open"`, `"File->Save As"`, `"Tajima (*.dst)"` dropdown option, and `"File name:"` labels are correct.
  - pywinauto's `print_control_identifiers()` can dump the actual UI tree to find correct element names.
- **Wings auto-detection paths** — converter searches 8 hardcoded paths + system PATH. If FM has Wings installed somewhere else, need to add that path to `_EDITOR_SEARCH` in `converter.py`.

### DECISIONS STILL OPEN:
- **Desktop vs web app** — both use identical core logic (same `app/core/*` modules). UI decision deferred until logic is tested on Windows with real files. Desktop app already built and launched; web demo also working.

## What To Do Next

### Step 1: Test on Windows
1. Copy project folder to Windows laptop
2. Install Python 3.12+: `python.org/downloads`
3. Install deps: `pip install pyembroidery olefile pywinauto customtkinter`
4. **First: verify converter works manually:**
   ```python
   from pywinauto import Application
   # Open My Editor, use print_control_identifiers() to dump UI tree
   # Compare against strings in converter.py — fix any mismatches
   ```
5. **Then: test full pipeline via desktop app:**
   ```
   python main.py
   # Browse to folder with NGS files → Combine → verify output in Wings
   ```

### Step 2: Build the .exe
On the Windows machine:
```
pip install nuitka
python build.py
```
This creates `EmbroideryC.exe` — a single standalone file. Operator double-clicks, no Python needed.

### Step 3: Deploy updates
If a bug is found after shipping:
1. Fix the code
2. Re-run `python build.py` to get a new .exe
3. Replace the old .exe on the operator's machine
(Consider adding auto-update later if update frequency is high.)

## How to Run (Mac — development only)
```
cd /Users/shaan_pawa/Micro/embroidery-combiner
source .venv/bin/activate              # Python 3.12 venv (Tk 9.0)
python main.py                         # Desktop app (DST only, no NGS conversion)
python web_demo.py                     # Web demo at http://localhost:5123
python -m pytest tests/ -v             # 42 tests
```
**Mac setup note:** System Python 3.9 has Tk 8.5 (too old for customtkinter). Need Python 3.12 via Homebrew:
```
brew install python@3.12 python-tk@3.12
/opt/homebrew/bin/python3.12 -m venv .venv
.venv/bin/pip install pyembroidery olefile customtkinter flask
```

## How to Run (Windows — production)
```
cd embroidery-combiner
pip install pyembroidery olefile pywinauto customtkinter
python main.py
```

## Edge Cases Handled
- Previous combined output in folder (e.g., 216-225.dst): auto-skipped by `is_range_filename()`
- Mixed folder (NGS + DST): converts NGS, combines all (but same-number NGS+DST would duplicate — unlikely in practice)
- Conversion failure on one file: continues with others, reports which failed
- Non-embroidery files: silently ignored
- Empty folder / single file: clear warning messages
- Sequence gaps and duplicate numbers: warned

## Files to Ignore
Legacy duplicates at project root (superseded by `app/` package):
- `combiner.py`, `converter.py`, `validator.py`, `licensing.py`, `gui.py`
- These are older versions, NOT used by the app. The real code is in `app/core/` and `app/`.
