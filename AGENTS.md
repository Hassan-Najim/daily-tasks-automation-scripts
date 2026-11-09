# AGENTS.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview

A collection of Python automation scripts organized by category, distributed as a
Textual TUI Windows executable (`daily-tasks.exe`) via a zero-install PowerShell
launcher (`launcher.ps1`).

## Project Structure

```
├── main.py                       # Plain CLI menu (loads scripts/ dynamically)
├── main_bundled.py               # Textual TUI + PyInstaller build entry (registry/probes here)
├── script_common.py              # Shared script helpers (ext lists, natural sort, HEIC, Office export)
├── launcher.ps1                  # irm|iex launcher: download/verify/run + -Install/-Uninstall
├── .github/workflows/release.yml # Tag v* -> build exe -> GitHub Release + SHA256SUMS
├── scripts/
│   ├── pdf/                      # everything_to_pdf, pdf_merge_split, images_to_pdf,
│   │                             #   images_to_pdf_custom_aspect, word_to_pdf, pdf_to_md
│   ├── image/                    # image_format_converter, image_resizer, phone_frame_overlay
│   └── files/                    # file_organizer, batch_renamer
└── requirements.txt              # All dependencies (textual, Pillow, pillow-heif, reportlab,
                                  #   pypdf, pdfplumber, pywin32)
```

## Running

```bash
python main.py            # plain menu (operates on the launch folder)
python main_bundled.py    # Textual TUI (working directory, D to change)
python main_bundled.py --build   # PyInstaller exe -> dist/daily-tasks.exe
```

End user: `irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1 | iex`
or the installed `daily-tasks` command (shim in `%LOCALAPPDATA%\daily-tasks`, on
user PATH; the exe is cached in the `app\` subfolder so it cannot shadow the
shim — PATHEXT resolves `.exe` before `.cmd`; an HKCU App Paths entry maps the
bare name to the shim for Win+R and the Explorer address bar).

## Key Patterns

- `scripts/<category>/<name>/script.py` is the **single source of truth** — there are
  no embedded copies. Both `main.py` and `main_bundled.py` load these files dynamically
  (`importlib`; in the frozen exe they ship as PyInstaller data via `--add-data` and are
  found under `sys._MEIPASS/scripts`). `main_bundled.py` holds only TUI metadata
  (ScriptEntry, probes) and the build function.
- Canonical script contract: every `script.py` exposes `main(log=print, ask=None)` and
  operates on `os.getcwd()`. `log(str)` is the output sink; `ask(prompt, options) ->
  str|None` asks a multiple-choice pick (TUI modal; console fallback = numbered
  `input()` — no free-text entry exists in the TUI sink). Scripts must also work when
  run directly (`python script.py`) and import `script_common` with the repo-root
  fallback bootstrap.
- Heavy libraries (PIL, reportlab, pdfplumber, pypdf, docx, win32com) are imported
  lazily inside functions so the TUI starts fast and imports happen in the worker thread.
- `script_common.py` centralizes shared behavior: extension constants, `natural_key`,
  `register_heif()` (optional pillow-heif), `office_export_pdf()` (COM export with
  retry; used by word_to_pdf and everything_to_pdf).
- ScriptEntry uses a unified detail format: `what` (one sentence), `needs`
  (requirements list), `how` (numbered steps), `output` (what is created),
  and `probe` (live readiness line for the detail pane; "Ready: ..." is green).
- The TUI runs scripts in thread workers via `@work(thread=True)` and talks to the
  UI with `app.call_from_thread`; blocking waits use `threading.Event`
- `launcher.ps1` must stay PowerShell 5.1 compatible (no ternary/`??` operators)
- Release asset names are load-bearing: `daily-tasks.exe` and `SHA256SUMS.txt`
  are hardcoded in `launcher.ps1`
- PyInstaller build flags in `build_exe()` are load-bearing: `--add-data scripts`
  ships the script files, and every library the scripts import must be listed via
  `--collect-all` because data-file scripts are not traced by static analysis.

## Adding New Scripts

1. Create folder under appropriate category: `scripts/<category>/<script_name>/`
   with `script.py` (`def main(log=print, ask=None)`), `README.md`, `requirements.txt`
2. Register in `main.py` SCRIPTS dict AND in `main_bundled.py`
   (`_SCRIPT_FILES` path + `ScriptEntry` with `what/needs/how/output/probe`)
3. Update the scripts table in `README.md`
4. Bump `__version__` in `main_bundled.py`, tag `v*` to release

## Releasing

1. Update `__version__` in `main_bundled.py` and commit
2. `git tag vX.Y.Z && git push origin vX.Y.Z`
3. The workflow builds and publishes the release; the launcher auto-updates users

## Dependencies

Install all: `pip install -r requirements.txt`

| Package | Used By |
|---------|---------|
| textual | main_bundled.py TUI |
| Pillow | All image scripts |
| pillow-heif | HEIC support (optional at runtime; bundled in the exe) |
| reportlab | PDF-producing scripts |
| pypdf | pdf_merge_split, everything_to_pdf |
| pdfplumber | pdf_to_md |
| pywin32 | Office COM exports (Windows only) |
