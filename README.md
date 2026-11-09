# Daily Tasks — Automation Suite

A Textual TUI bundling everyday-task automations: combine anything into a
PDF, merge/split PDFs, images to PDF, Word to PDF, PDF to Markdown,
resizing, format conversion, phone-screenshot mockups, file organizing and
batch renaming — as a single portable Windows executable.

## Run it now (zero install)

Open **PowerShell** on Windows 10/11 and paste:

```powershell
irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1 | iex
```

**Using cmd.exe instead?** Either type `powershell` and press Enter first, or paste this directly:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1 | iex"
```

The launcher downloads the latest release (SHA256-verified), caches it in
`%LOCALAPPDATA%\daily-tasks`, and starts the TUI. Re-running the command
auto-updates to the newest release. Offline? The cached copy is used.

## Install the `daily-tasks` command

Run the launcher once and answer **y** at the prompt — or install directly:

```powershell
iex "& { $(irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1) } -Install"
```

From **cmd.exe**:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1 -OutFile $env:TEMP\daily-tasks.ps1; & \"$env:TEMP\daily-tasks.ps1\" -Install"
```

This drops a `daily-tasks` command into your user PATH (works in cmd,
PowerShell, and Win+R). Open a **new** terminal and type:

```
daily-tasks
```

Uninstall anytime:

```powershell
iex "& { $(irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1) } -Uninstall"
```

From **cmd.exe**:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1 -OutFile $env:TEMP\daily-tasks.ps1; & \"$env:TEMP\daily-tasks.ps1\" -Uninstall"
```

## Scripts

| Category | Script | Description |
|----------|--------|-------------|
| [PDF](scripts/pdf/everything_to_pdf) | Everything to PDF | Combine Word, Excel, PowerPoint, PDF and images into one PDF |
| [PDF](scripts/pdf/pdf_merge_split) | PDF Merge & Split | Merge all PDFs into one, or split one into pages |
| [PDF](scripts/pdf/images_to_pdf) | Images to PDF | Convert images to a single PDF (EXIF-aware) |
| [PDF](scripts/pdf/images_to_pdf_custom_aspect) | Images to PDF (Custom) | Convert with native aspect ratios |
| [PDF](scripts/pdf/word_to_pdf) | Word to PDF | Batch convert Word docs to PDF |
| [PDF](scripts/pdf/pdf_to_md) | PDF to Markdown | Convert text-based PDFs to Markdown (no AI/OCR) |
| [Image](scripts/image/image_format_converter) | Image Converter | Convert images between formats (incl. HEIC via pillow-heif) |
| [Image](scripts/image/image_resizer) | Image Resizer | Batch resize or compress images to a target size |
| [Image](scripts/image/phone_frame_overlay) | Phone Frame Overlay | Add device frames to screenshots |
| [Files](scripts/files/file_organizer) | File Organizer | Sort loose files into subfolders by type or year-month |
| [Files](scripts/files/batch_renamer) | Batch Renamer | Rename files with one rule: numbered, lowercase-dashes, or date-stamped |

Each script operates on the TUI's working directory (change it with `D`).
Selecting a script shows a unified detail view — **WHAT** it does, what it
**NEEDS**, step-by-step **HOW** to use it, what it **OUTPUT**s — plus a live
check of whether the current folder has matching files to work with.
Press `F` to filter the script list.

## TUI keys

| Key | Action |
|-----|--------|
| `↑` `↓` | Navigate scripts (themes preview live in the theme picker) |
| `Enter` / `R` | Run selected script |
| `F` | Filter the script list (Esc clears and returns to the list) |
| `D` | Change working directory (remembered between launches) |
| `O` | Open the working folder in Explorer |
| `T` | Theme picker (scroll to preview, Enter to keep, Esc to cancel) |
| `Ctrl+P` | Command palette |
| `C` | Clear the log panel |
| `Q` | Quit |

Your theme and working directory are saved in
`%LOCALAPPDATA%\daily-tasks\settings.json`.

## Requirements

- Windows 10/11 for the launcher/executable (no Python needed)
- Microsoft Word for **Word to PDF**; Microsoft Office (Word/Excel/
  PowerPoint) for the matching file types in **Everything to PDF** —
  both use COM automation and skip files whose app is missing

## Development

```bash
pip install -r requirements.txt

# Run the TUI from source
python main_bundled.py

# Run the plain CLI menu instead
python main.py

# Build the standalone executable locally
python main_bundled.py --build
```

The `scripts/` folders are the single source of truth: both `main.py` and
the TUI load `scripts/<category>/<name>/script.py` at runtime (in the exe
they ship as data files), and each script exposes
`main(log=print, ask=None)` and operates on the current working directory.

### Releases

Tag a version (`v1.0.0`) and push — the [Release workflow](.github/workflows/release.yml)
builds `daily-tasks.exe` on a Windows runner, generates `SHA256SUMS.txt`, and
publishes a GitHub Release. The launcher picks up new releases automatically.

### Security note

`irm ... | iex` executes remote code — only ever run the one-liner from this
README. The launcher verifies each download against the SHA256 checksum
published in the same release, and all transport is HTTPS.

## Project Structure

```
daily-tasks-automation-scripts/
├── main.py                   # Plain CLI menu (loads scripts/ dynamically)
├── main_bundled.py           # Textual TUI + PyInstaller build entry
├── script_common.py          # Shared script helpers (HEIC, sorting, Office export)
├── launcher.ps1              # Zero-install launcher + command installer
├── requirements.txt          # All dependencies
├── .github/workflows/release.yml
└── scripts/                  # Single source of truth, loaded by both entry points
    ├── pdf/                  #   script.py + README.md + requirements.txt
    ├── image/
    └── files/
```
