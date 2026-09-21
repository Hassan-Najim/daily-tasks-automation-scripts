# Daily Tasks — Automation Suite

A Textual TUI bundling everyday-task automations: images to PDF, Word to PDF,
phone-screenshot mockups, format conversion, and bulk find & replace — as a
single portable Windows executable.

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
| [PDF](scripts/pdf/images_to_pdf) | Images to PDF | Convert images to a single PDF |
| [PDF](scripts/pdf/images_to_pdf_custom_aspect) | Images to PDF (Custom) | Convert with native aspect ratios |
| [PDF](scripts/pdf/word_to_pdf) | Word to PDF | Batch convert Word docs to PDF |
| [Image](scripts/image/phone_frame_overlay) | Phone Frame Overlay | Add device frames to screenshots |
| [Image](scripts/image/image_format_converter) | Image Converter | Batch convert images to PNG |
| [Document](scripts/document/word_find_replace) | Word Find & Replace | Bulk text replacement in Word docs |

Each script operates on the TUI's working directory (change it with `D`).

## TUI keys

| Key | Action |
|-----|--------|
| `↑` `↓` | Navigate scripts |
| `Enter` / `R` | Run selected script |
| `D` | Change working directory |
| `C` | Clear the log panel |
| `Q` | Quit |

## Requirements

- Windows 10/11 for the launcher/executable (no Python needed)
- Microsoft Word installed for **Word to PDF** (uses COM automation)

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
├── main.py                   # Plain CLI menu (runs scripts/ standalone)
├── main_bundled.py           # Textual TUI + PyInstaller build entry
├── launcher.ps1              # Zero-install launcher + command installer
├── requirements.txt          # All dependencies
├── .github/workflows/release.yml
└── scripts/                  # Standalone per-script folders
    ├── pdf/                  #   script.py + README.md + requirements.txt
    ├── image/
    └── document/
```
