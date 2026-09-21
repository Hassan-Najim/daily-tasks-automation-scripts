"""
Daily Tasks - Automation Suite (Bundled Version)
All scripts embedded, packaged as a Textual TUI for standalone
executable compilation.

Run the TUI:       python main_bundled.py
Build the exe:     python main_bundled.py --build
"""

import os
import re
import sys
import json
import time
import threading
from dataclasses import dataclass
from pathlib import Path

from rich.text import Text

try:
    from textual import work
    from textual import on
    from textual.app import App, ComposeResult
    from textual.binding import Binding
    from textual.containers import Horizontal, Vertical
    from textual.screen import ModalScreen
    from textual.widgets import Footer, Header, Input, Label, ListItem, ListView, RichLog, Static
except ImportError:
    print("Error: textual is not installed. Run: pip install textual")
    sys.exit(1)


__version__ = "1.1.0"

APP_NAME = "daily-tasks"


# =============================================================================
# EMBEDDED SCRIPTS
# =============================================================================
# Every script accepts:
#   log(prompt)            -> output sink (default: print)
#   ask(prompt, options)   -> ask the user to pick one option, returns the
#                             chosen string or None if cancelled


def images_to_pdf_main(log=print, ask=None):
    from PIL import Image
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter

    folder_path = os.getcwd()
    output_pdf = os.path.join(folder_path, "output.pdf")

    extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
    images = [f for f in os.listdir(folder_path) if f.lower().endswith(extensions)]
    images.sort()

    if not images:
        log("No images found in the current folder.")
        return

    c = canvas.Canvas(output_pdf, pagesize=letter)
    width, height = letter

    for img in images:
        img_path = os.path.join(folder_path, img)
        log(f"Processing: {img}")

        image = Image.open(img_path)
        aspect = image.width / float(image.height)

        if aspect > 1:
            img_width = width
            img_height = width / aspect
        else:
            img_height = height
            img_width = height * aspect

        x = (width - img_width) / 2
        y = (height - img_height) / 2

        c.drawImage(img_path, x, y, img_width, img_height)
        c.showPage()

    c.save()
    log(f"\nPDF saved: {output_pdf}")


def images_to_pdf_custom_main(log=print, ask=None):
    from PIL import Image, ImageOps
    from reportlab.pdfgen import canvas

    def natural_key(s):
        return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]

    folder_path = os.getcwd()
    output_pdf = os.path.join(folder_path, "output.pdf")
    default_dpi = 72

    extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff')
    images = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
              if f.lower().endswith(extensions)]

    if not images:
        log("No images found in the current folder.")
        return

    images.sort(key=lambda p: natural_key(os.path.basename(p)))
    c = canvas.Canvas(output_pdf)

    for img_path in images:
        log(f"Processing: {os.path.basename(img_path)}")

        with Image.open(img_path) as im:
            im = ImageOps.exif_transpose(im)
            px_w, px_h = im.size
            dpi_x, dpi_y = im.info.get("dpi", (default_dpi, default_dpi))

            page_w = px_w * 72.0 / (dpi_x or default_dpi)
            page_h = px_h * 72.0 / (dpi_y or default_dpi)

            c.setPageSize((page_w, page_h))
            c.drawImage(img_path, 0, 0, width=page_w, height=page_h,
                        preserveAspectRatio=False, mask='auto')
            c.showPage()

    c.save()
    log(f"\nPDF saved: {output_pdf}")


def word_to_pdf_main(log=print, ask=None):
    if sys.platform != "win32":
        log("Error: Word to PDF only works on Windows with Microsoft Word installed.")
        return

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        log("Error: pywin32 not installed. Run: pip install pywin32")
        return

    folder = Path(os.getcwd())

    files = [
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in (".docx", ".doc")
        and not p.name.startswith("~$")
    ]

    if not files:
        log("No .doc or .docx files found in current folder.")
        return

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.ScreenUpdating = False
    word.DisplayAlerts = 0

    failures = []
    converted = 0

    try:
        for doc_path in files:
            pdf_path = doc_path.with_suffix(".pdf")

            if pdf_path.exists():
                log(f"Skipping (exists): {doc_path.name}")
                continue

            log(f"Converting: {doc_path.name}")

            success = False
            for attempt in range(1, 4):
                try:
                    doc = word.Documents.Open(str(doc_path))
                    doc.ExportAsFixedFormat(
                        OutputFileName=str(pdf_path),
                        ExportFormat=17,
                        OpenAfterExport=False,
                        OptimizeFor=0,
                        Range=0,
                        From=1, To=1,
                        Item=0,
                        IncludeDocProps=True,
                        KeepIRM=True,
                        CreateBookmarks=1,
                        DocStructureTags=True,
                        BitmapMissingFonts=True,
                        UseISO19005_1=False
                    )
                    doc.Close(False)
                    success = True
                    converted += 1
                    break
                except Exception as e:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                    if attempt < 3:
                        time.sleep(1.2 * attempt)
                    else:
                        failures.append((doc_path.name, str(e)))

            if not success:
                log(f"  FAILED: {doc_path.name}")
    finally:
        word.Quit()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    log(f"\nConverted: {converted}")
    if failures:
        log(f"Failed: {len(failures)}")
        for name, err in failures:
            log(f"  - {name}: {err}")


def phone_frame_overlay_main(log=print, ask=None):
    from PIL import Image

    folder = os.getcwd()
    frame_path = os.path.join(folder, 'frame.png')
    screens_folder = os.path.join(folder, 'screens')
    output_folder = os.path.join(folder, 'output')

    if not os.path.exists(frame_path):
        log(f"Error: frame.png not found in {folder}")
        log("Please add a frame.png file (device mockup with transparency)")
        return

    if not os.path.exists(screens_folder):
        log(f"Error: 'screens' folder not found in {folder}")
        log("Please create a 'screens' folder and add your screenshots")
        return

    os.makedirs(output_folder, exist_ok=True)

    frame = Image.open(frame_path).convert('RGBA')
    frame_width, frame_height = frame.size

    processed = 0
    for filename in os.listdir(screens_folder):
        if filename.lower().endswith('.png') and not filename.startswith('.'):
            screen_path = os.path.join(screens_folder, filename)

            try:
                screen = Image.open(screen_path).convert('RGBA')
            except Exception as e:
                log(f"Skipping {filename}: {e}")
                continue

            canvas = Image.new('RGBA', (frame_width, frame_height), (0, 0, 0, 0))
            screen_x = (frame_width - screen.width) // 2
            screen_y = (frame_height - screen.height) // 2
            canvas.paste(screen, (screen_x, screen_y), screen)

            final_image = Image.alpha_composite(canvas, frame)

            output_path = os.path.join(output_folder, filename)
            final_image.save(output_path, format='PNG')
            log(f"Created: {filename}")
            processed += 1

    log(f"\nProcessed {processed} images -> output/")


def image_format_converter_main(log=print, ask=None):
    from PIL import Image

    folder_path = os.getcwd()
    supported = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff', '.webp')
    target_formats = ['PNG', 'JPG', 'BMP', 'GIF', 'TIFF', 'WEBP']
    exts_by_format = {'PNG': ('.png',), 'JPG': ('.jpg', '.jpeg'), 'BMP': ('.bmp',),
                      'GIF': ('.gif',), 'TIFF': ('.tif', '.tiff'), 'WEBP': ('.webp',)}

    images = sorted(f for f in os.listdir(folder_path)
                    if f.lower().endswith(supported))
    if not images:
        log("No supported images found in the working directory.")
        return

    choice = None
    if ask:
        choice = ask("Select target format", target_formats)
    else:
        print("Select target format:")
        for i, fmt in enumerate(target_formats, 1):
            print(f"  {i}. {fmt}")
        raw = input(f"Choice [1-{len(target_formats)}]: ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(target_formats):
            choice = target_formats[int(raw) - 1]
    if not choice:
        log("Cancelled.")
        return

    save_format = 'JPEG' if choice == 'JPG' else choice
    target_ext = exts_by_format[choice][0]
    log(f"Converting {len(images)} image(s) to {choice} (originals are kept)\n")

    converted = 0
    skipped = 0
    failed = 0
    for filename in images:
        if filename.lower().endswith(exts_by_format[choice]):
            skipped += 1
            continue
        file_path = os.path.join(folder_path, filename)
        new_filename = os.path.splitext(filename)[0] + target_ext
        new_filepath = os.path.join(folder_path, new_filename)

        try:
            with Image.open(file_path) as img:
                img.load()
                if choice == 'JPG':
                    if img.mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGBA')
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        background.paste(img, mask=img.split()[-1])
                        img = background
                    elif img.mode != 'RGB':
                        img = img.convert('RGB')
                elif choice == 'GIF':
                    if img.mode not in ('P', 'L', 'RGB'):
                        img = img.convert('P')
                else:
                    if img.mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGBA')
                    elif img.mode != 'RGB':
                        img = img.convert('RGB')
                img.save(new_filepath, save_format)

            log(f"Converted: {filename} -> {new_filename}")
            converted += 1
        except Exception as e:
            log(f"Failed: {filename}: {e}")
            failed += 1

    log(f"\nConverted {converted} image(s) to {choice}")
    if skipped:
        log(f"Skipped {skipped} already in target format")
    if failed:
        log(f"Failed: {failed}")


def word_find_replace_main(log=print, ask=None):
    from docx import Document

    folder = os.getcwd()

    json_path = os.path.join(folder, "corrections.json")
    if not os.path.exists(json_path):
        log("Error: corrections.json not found in current folder")
        log("Create a corrections.json file with format:")
        log('[{"OgSentence": "original text", "NewSentence": "replacement"}, ...]')
        return

    docx_files = [f for f in os.listdir(folder)
                  if f.lower().endswith('.docx') and not f.startswith('~$')]
    if not docx_files:
        log("Error: No .docx file found in current folder")
        return

    if len(docx_files) > 1:
        if ask is None:
            log("Multiple .docx files found:")
            for i, f in enumerate(docx_files, 1):
                log(f"  {i}. {f}")
            choice = input("Enter number: ").strip()
            try:
                doc_file = docx_files[int(choice) - 1]
            except (ValueError, IndexError):
                log("Invalid choice")
                return
        else:
            doc_file = ask("Multiple .docx files found - select one", docx_files)
            if doc_file is None:
                log("Cancelled.")
                return
    else:
        doc_file = docx_files[0]

    doc_path = os.path.join(folder, doc_file)

    with open(json_path, 'r', encoding='utf-8') as f:
        corrections = json.load(f)

    log(f"Document: {doc_file}")
    log(f"Corrections: {len(corrections)}")

    document = Document(doc_path)
    replacements = 0

    for entry in corrections:
        original = entry.get("OgSentence")
        replacement = entry.get("NewSentence")

        if not original or not replacement:
            continue

        for paragraph in document.paragraphs:
            if original in paragraph.text:
                paragraph.text = paragraph.text.replace(original, replacement, 1)
                replacements += 1
                break

        for table in document.tables:
            found = False
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if original in paragraph.text:
                            paragraph.text = paragraph.text.replace(original, replacement, 1)
                            replacements += 1
                            found = True
                            break
                    if found:
                        break
                if found:
                    break
            if found:
                break

    output_path = doc_path.replace(".docx", "_updated.docx")
    document.save(output_path)
    log(f"\nReplacements: {replacements}")
    log(f"Saved: {output_path}")


# =============================================================================
# SCRIPT REGISTRY
# =============================================================================

@dataclass
class ScriptEntry:
    name: str
    category: str
    description: str
    notes: str
    func: object


SCRIPTS = [
    ScriptEntry(
        name="Images to PDF",
        category="PDF",
        description="Combine every image in the working directory into a single "
                    "output.pdf, centered on US Letter pages.",
        notes="Supports PNG, JPG, JPEG, BMP, GIF. Landscape images fill page width, "
              "portrait images fill page height.",
        func=images_to_pdf_main,
    ),
    ScriptEntry(
        name="Images to PDF (Custom Aspect)",
        category="PDF",
        description="Combine images into a single output.pdf where each page keeps "
                    "the image's native size and aspect ratio (DPI-aware).",
        notes="Supports PNG, JPG, JPEG, BMP, GIF, TIFF. Handles EXIF rotation and "
              "natural filename ordering (img2 before img10).",
        func=images_to_pdf_custom_main,
    ),
    ScriptEntry(
        name="Word to PDF",
        category="PDF",
        description="Batch convert every .doc/.docx file in the working directory "
                    "to PDF via Microsoft Word automation.",
        notes="Windows + Microsoft Word required. Existing PDFs are skipped; failed "
              "conversions are retried up to 3 times.",
        func=word_to_pdf_main,
    ),
    ScriptEntry(
        name="Phone Frame Overlay",
        category="Image",
        description="Composite device frame mockups over your screenshots.",
        notes="Needs frame.png (transparent device frame) and a screens/ folder with "
              "PNG screenshots. Results go to output/.",
        func=phone_frame_overlay_main,
    ),
    ScriptEntry(
        name="Image Format Converter",
        category="Image",
        description="Batch convert all images in the working directory to a chosen "
                    "format: PNG, JPG, BMP, GIF, TIFF or WEBP.",
        notes="Accepts PNG, JPG, JPEG, BMP, GIF, TIF, TIFF, WEBP input. Transparent "
              "images are flattened onto white for JPG. Files already in the target "
              "format are skipped; originals are kept.",
        func=image_format_converter_main,
    ),
    ScriptEntry(
        name="Word Find & Replace",
        category="Document",
        description="Apply bulk sentence-level replacements to a Word document from "
                    "a corrections.json list.",
        notes="Needs corrections.json and at least one .docx in the working "
              "directory. Result saved as *_updated.docx.",
        func=word_find_replace_main,
    ),
]


# =============================================================================
# MODAL SCREENS
# =============================================================================

class InputModal(ModalScreen):
    """Text input dialog. Dismisses with the entered string, or None on Escape."""

    def __init__(self, prompt: str, initial: str = "") -> None:
        super().__init__()
        self.prompt = prompt
        self.initial = initial

    def compose(self) -> ComposeResult:
        yield Vertical(
            Label(self.prompt, id="input-modal-prompt"),
            Input(value=self.initial, id="input-modal-field"),
        )

    def on_mount(self) -> None:
        self.query_one("#input-modal-field", Input).focus()

    @on(Input.Submitted)
    def submitted(self, event: Input.Submitted) -> None:
        self.dismiss(event.value)

    def on_key(self, event) -> None:
        if event.key == "escape":
            self.dismiss(None)


class SelectModal(ModalScreen):
    """Option picker dialog. Dismisses with the chosen string, or None on Escape."""

    def __init__(self, prompt: str, options: list) -> None:
        super().__init__()
        self.prompt = prompt
        self.options = options
        self._by_item = {}

    def compose(self) -> ComposeResult:
        items = [ListItem(Label(str(option))) for option in self.options]
        yield Vertical(
            Label(self.prompt, id="select-modal-prompt"),
            ListView(*items, id="select-modal-list"),
        )

    def on_mount(self) -> None:
        list_view = self.query_one("#select-modal-list", ListView)
        for item, option in zip(list_view.children, self.options):
            self._by_item[item] = str(option)
        list_view.focus()

    @on(ListView.Selected)
    def selected(self, event: ListView.Selected) -> None:
        self.dismiss(self._by_item.get(event.item))

    def on_key(self, event) -> None:
        if event.key == "escape":
            self.dismiss(None)


# =============================================================================
# TUI APPLICATION
# =============================================================================

class AutomationApp(App):
    TITLE = "Daily Tasks - Automation Suite"
    SUB_TITLE = f"v{__version__}"

    CSS = """
    Screen {
        layout: vertical;
    }
    #body {
        height: 1fr;
    }
    #sidebar {
        width: 44;
        border-right: solid $accent;
        padding: 1 0;
    }
    #sidebar-header {
        padding: 0 2;
        color: $text-muted;
        text-style: bold;
    }
    #script-list {
        height: 1fr;
        padding: 0 1;
    }
    #main {
        padding: 1 2;
    }
    #script-title {
        text-style: bold;
        color: $text;
        width: 1fr;
    }
    #script-category {
        color: $accent;
    }
    #script-description {
        margin-top: 1;
        width: 1fr;
    }
    #script-notes {
        margin-top: 1;
        width: 1fr;
        color: $text-muted;
    }
    #log-panel {
        height: 16;
        border-top: solid $accent;
    }
    InputModal Vertical {
        height: auto;
        margin: 4 8;
        padding: 1 2;
        background: $surface;
        border: round $accent;
    }
    #input-modal-prompt {
        width: 1fr;
    }
    SelectModal Vertical {
        height: auto;
        max-height: 60%;
        margin: 2 8;
        padding: 1 1;
        background: $surface;
        border: round $accent;
    }
    #select-modal-prompt {
        padding: 0 1;
    }
    #select-modal-list {
        height: auto;
        max-height: 20;
    }
    """

    BINDINGS = [
        Binding("r", "run_highlighted", "Run"),
        Binding("d", "change_directory", "Directory"),
        Binding("c", "clear_log", "Clear log"),
        Binding("q", "quit", "Quit"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self._entries = {}

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True)
        with Horizontal(id="body"):
            with Vertical(id="sidebar"):
                yield Label("SCRIPTS", id="sidebar-header")
                items = [ListItem(Label(script.name)) for script in SCRIPTS]
                yield ListView(*items, id="script-list", initial_index=0)
            with Vertical(id="main"):
                yield Label("", id="script-title")
                yield Label("", id="script-category")
                yield Static("", id="script-description")
                yield Static("", id="script-notes")
        yield RichLog(id="log-panel", markup=False, highlight=False, wrap=True)
        yield Footer()

    def on_mount(self) -> None:
        list_view = self.query_one("#script-list", ListView)
        for item, script in zip(list_view.children, SCRIPTS):
            self._entries[item] = script
        self._refresh_directory()
        self._show_script(SCRIPTS[0])
        list_view.focus()
        self.log_line(f"Daily Tasks v{__version__} - ready.")
        self.log_markup("Select a script, press [bold]Enter[/] or [bold]R[/] to run, "
                        "[bold]D[/] to change the working directory.")

    def _refresh_directory(self) -> None:
        self.sub_title = f"v{__version__} - {os.getcwd()}"

    def _show_script(self, script: ScriptEntry) -> None:
        self.query_one("#script-title", Label).update(script.name)
        self.query_one("#script-category", Label).update(f"[{script.category}]")
        self.query_one("#script-description", Static).update(script.description)
        self.query_one("#script-notes", Static).update(f"Note: {script.notes}")

    def _current_script(self):
        list_view = self.query_one("#script-list", ListView)
        if list_view.index is None:
            return None
        return SCRIPTS[list_view.index]

    def log_line(self, line: str) -> None:
        self.query_one("#log-panel", RichLog).write(line)

    def log_markup(self, markup: str) -> None:
        self.query_one("#log-panel", RichLog).write(Text.from_markup(markup))

    @on(ListView.Highlighted)
    def script_highlighted(self, event: ListView.Highlighted) -> None:
        if event.item is not None and event.item in self._entries:
            self._show_script(self._entries[event.item])

    @on(ListView.Selected)
    def script_selected(self, event: ListView.Selected) -> None:
        if event.item is not None and event.item in self._entries:
            self._run_script(self._entries[event.item])

    def action_run_highlighted(self) -> None:
        script = self._current_script()
        if script is not None:
            self._run_script(script)

    def action_clear_log(self) -> None:
        self.query_one("#log-panel", RichLog).clear()

    def action_change_directory(self) -> None:
        self.push_screen(InputModal("Working directory:", initial=os.getcwd()),
                         callback=self._directory_chosen)

    def _directory_chosen(self, choice) -> None:
        if not choice:
            return
        path = os.path.expandvars(os.path.expanduser(choice.strip().strip('"')))
        if os.path.isdir(path):
            os.chdir(path)
            self._refresh_directory()
            self.log_line(f"Working directory changed to: {os.getcwd()}")
            self.notify(f"Working directory: {os.getcwd()}")
        else:
            self.log_line(f"Not a directory: {path}")
            self.notify(f"Not a directory: {path}", severity="error")

    def _push_ask(self, prompt: str, options: list, on_result) -> None:
        self.push_screen(SelectModal(prompt, options), callback=on_result)

    def _ask_sync(self, prompt: str, options: list):
        holder = {}
        done = threading.Event()

        def on_result(choice):
            holder["choice"] = choice
            done.set()

        self.app.call_from_thread(self._push_ask, prompt, options, on_result)
        done.wait()
        return holder.get("choice")

    @work(thread=True, exclusive=True, group="script")
    def _run_script(self, script: ScriptEntry) -> None:
        def sink(line: str) -> None:
            self.app.call_from_thread(self.log_line, str(line))

        def ask(prompt: str, options: list):
            return self._ask_sync(prompt, options)

        self.app.call_from_thread(self.log_line, "")
        self.app.call_from_thread(self.log_markup, f"[bold]{'=' * 46}[/]")
        self.app.call_from_thread(self.log_markup, f"[bold]Running: {script.name}[/]")
        self.app.call_from_thread(self.log_markup, f"[bold]Directory: {os.getcwd()}[/]")
        self.app.call_from_thread(self.log_markup, f"[bold]{'=' * 46}[/]")

        started = time.monotonic()
        try:
            script.func(log=sink, ask=ask)
        except Exception as e:
            self.app.call_from_thread(self.log_line, f"Error: {e}")
            self.app.call_from_thread(
                self.notify, f"{script.name} failed: {e}", severity="error")
        else:
            elapsed = time.monotonic() - started
            self.app.call_from_thread(
                self.log_line, f"Done in {elapsed:.1f}s.")
            self.app.call_from_thread(
                self.notify, f"{script.name} finished",
                severity="information")


def main():
    app = AutomationApp()
    app.run()


# =============================================================================
# BUILD FUNCTION
# =============================================================================

def build_exe():
    """Compile this script into a standalone executable."""
    import subprocess

    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])

    script_path = Path(__file__).resolve()
    base_dir = script_path.parent

    print("=" * 50)
    print("Building Daily Tasks Executable")
    print("=" * 50)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--name", APP_NAME,
        "--console",
        "--clean",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--hidden-import", "PIL.ImageOps",
        "--hidden-import", "reportlab",
        "--hidden-import", "reportlab.pdfgen",
        "--hidden-import", "reportlab.pdfgen.canvas",
        "--hidden-import", "reportlab.lib.pagesizes",
        "--hidden-import", "docx",
        "--hidden-import", "docx.document",
        "--collect-submodules", "reportlab",
        "--collect-submodules", "docx",
        "--collect-all", "textual",
        "--collect-all", "rich",
        str(script_path)
    ]

    if sys.platform == "win32":
        cmd.extend([
            "--hidden-import", "win32com",
            "--hidden-import", "win32com.client",
            "--hidden-import", "pythoncom",
        ])

    print("\nRunning PyInstaller...\n")
    result = subprocess.run(cmd, cwd=base_dir)

    if result.returncode == 0:
        exe_name = f"{APP_NAME}.exe" if sys.platform == "win32" else APP_NAME
        exe_path = base_dir / "dist" / exe_name
        print("\n" + "=" * 50)
        print("BUILD SUCCESSFUL!")
        print("=" * 50)
        print(f"\nExecutable: {exe_path}")
        if exe_path.exists():
            print(f"Size: {exe_path.stat().st_size / 1024 / 1024:.1f} MB")
        print("\nCopy this file anywhere and run it.")
    else:
        print("\nBuild failed!")
        sys.exit(1)


if __name__ == "__main__":
    if "--build" in sys.argv:
        build_exe()
    else:
        main()
