# AGENTS.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview

A collection of Python automation scripts organized by category, distributed as a
Textual TUI Windows executable (`daily-tasks.exe`) via a zero-install PowerShell
launcher (`launcher.ps1`).

## Project Structure

```
├── main.py                       # Plain CLI menu (imports scripts/ standalone)
├── main_bundled.py               # Textual TUI with all scripts embedded; exe entry
├── launcher.ps1                  # irm|iex launcher: download/verify/run + -Install/-Uninstall
├── .github/workflows/release.yml # Tag v* -> build exe -> GitHub Release + SHA256SUMS
├── scripts/
│   ├── pdf/                      # PDF scripts (images_to_pdf, custom_aspect, word_to_pdf)
│   ├── image/                    # Image scripts (phone_frame_overlay, image_format_converter)
│   └── document/                 # Document scripts (word_find_replace)
└── requirements.txt              # All dependencies (textual, Pillow, reportlab, python-docx, pywin32)
```

## Running

```bash
python main.py            # plain menu
python main_bundled.py    # Textual TUI
python main_bundled.py --build   # PyInstaller exe -> dist/daily-tasks.exe
```

End user: `irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1 | iex`
or the installed `daily-tasks` command (shim in `%LOCALAPPDATA%\daily-tasks`, on user PATH).

## Key Patterns

- Standalone scripts live in `scripts/<category>/<name>/` with `script.py` (entry
  point `main()`), `README.md`, `requirements.txt`
- Embedded scripts in `main_bundled.py` accept `(log=print, ask=None)` sinks:
  `log(str)` for output, `ask(prompt, options) -> str|None` for interactive picks
  (TUI shows a modal; console falls back to numbered input)
- The TUI runs scripts in thread workers via `@work(thread=True)` and talks to the
  UI with `app.call_from_thread`; blocking waits use `threading.Event`
- `launcher.ps1` must stay PowerShell 5.1 compatible (no ternary/`??` operators)
- Release asset names are load-bearing: `daily-tasks.exe` and `SHA256SUMS.txt`
  are hardcoded in `launcher.ps1`

## Adding New Scripts

1. Create folder under appropriate category: `scripts/<category>/<script_name>/`
2. Add `script.py` with a `main()` function
3. Add `README.md` and `requirements.txt`
4. Register in `main.py` SCRIPTS dict AND in `main_bundled.py`:
   copy the function (with `log`/`ask` sinks) and add a `ScriptEntry` to SCRIPTS
5. Update the scripts table in `README.md`
6. Bump `__version__` in `main_bundled.py`, tag `v*` to release

## Releasing

1. Update `__version__` in `main_bundled.py` and commit
2. `git tag vX.Y.Z && git push origin vX.Y.Z`
3. The workflow builds and publishes the release; the launcher auto-updates users

## Dependencies

Install all: `pip install -r requirements.txt`

| Package | Used By |
|---------|--------|
| textual | main_bundled.py TUI |
| Pillow | All image scripts |
| reportlab | PDF scripts |
| python-docx | word_find_replace |
| pywin32 | word_to_pdf (Windows only) |
