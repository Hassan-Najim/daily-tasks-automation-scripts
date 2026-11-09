"""
Daily Tasks - Automation Suite
Textual TUI that loads the scripts from scripts/ (the single source of
truth) and packages them as a standalone Windows executable.

Run the TUI:       python main_bundled.py
Build the exe:     python main_bundled.py --build
"""

import importlib.util
import os
import re
import sys
import json
import time
import threading
from dataclasses import dataclass
from pathlib import Path

from rich.text import Text

import script_common as sc

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


__version__ = "1.3.1"

APP_NAME = "daily-tasks"

DEFAULT_THEME = "textual-dark"


def _settings_file() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    if not base:
        base = str(Path.home() / f".{APP_NAME}")
    return Path(base) / APP_NAME / "settings.json"


def load_settings() -> dict:
    try:
        with open(_settings_file(), "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(settings: dict) -> None:
    try:
        path = _settings_file()
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except OSError:
        pass


def update_setting(key: str, value) -> None:
    settings = load_settings()
    settings[key] = value
    save_settings(settings)


# =============================================================================
# SCRIPT LOADING
# =============================================================================
# The scripts in scripts/<category>/<name>/script.py are the single source
# of truth. They are loaded dynamically (in the frozen exe they ship as
# PyInstaller data files). Each script exposes main(log=print, ask=None)
# and operates on os.getcwd().
#
# Loading is deferred until first run so the TUI starts instantly and the
# heavy libraries (PIL, reportlab, pdfplumber, ...) import inside the run
# worker thread.

_SCRIPTS_DIR = "scripts"

# Display name -> relative path under scripts/
_SCRIPT_FILES = {
    "Images to PDF": "pdf/images_to_pdf/script.py",
    "Images to PDF (Custom Aspect)": "pdf/images_to_pdf_custom_aspect/script.py",
    "Word to PDF": "pdf/word_to_pdf/script.py",
    "PDF to Markdown": "pdf/pdf_to_md/script.py",
    "Image Format Converter": "image/image_format_converter/script.py",
    "Image Resizer": "image/image_resizer/script.py",
    "Phone Frame Overlay": "image/phone_frame_overlay/script.py",
    "File Organizer": "files/file_organizer/script.py",
    "Batch Renamer": "files/batch_renamer/script.py",
}


def _scripts_base_dir() -> Path:
    """Directory that contains scripts/ (dev tree or extracted exe data)."""
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass and (Path(meipass) / _SCRIPTS_DIR).is_dir():
            return Path(meipass)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _load_script_main(rel_path: str):
    """Return a main(log, ask) that imports the script file on first call."""
    base = _scripts_base_dir()
    path = base / _SCRIPTS_DIR / rel_path
    module_name = "dt_script_" + re.sub(r"\W+", "_", rel_path)
    cache = {}

    def main(log=print, ask=None):
        if "module" not in cache:
            if not path.is_file():
                raise FileNotFoundError(f"Script file not found: {path}")
            spec = importlib.util.spec_from_file_location(module_name, path)
            module = importlib.util.module_from_spec(spec)
            sys.modules[module_name] = module
            spec.loader.exec_module(module)
            cache["module"] = module
        return cache["module"].main(log=log, ask=ask)

    return main


_SCRIPT_FUNCS = {name: _load_script_main(rel)
                 for name, rel in _SCRIPT_FILES.items()}


# =============================================================================
# SCRIPT REGISTRY
# =============================================================================
# Every entry uses the same four sections:
#   what   -> one sentence: what the tool does functionally
#   needs  -> files / software / OS requirements
#   how    -> numbered usage steps
#   output -> what gets created and where
# `probe` returns a live status line for the current working directory.

_IMAGE_EXTS = sc.IMAGE_EXTS + sc.HEIC_EXTS


def _heif_available() -> bool:
    try:
        import pillow_heif  # noqa: F401
        return True
    except ImportError:
        return False


def _count_files(extensions, folder=None):
    folder = folder or os.getcwd()
    try:
        return sum(1 for f in os.listdir(folder)
                   if f.lower().endswith(extensions) and not f.startswith('~$'))
    except OSError:
        return 0


def _plural(count, word):
    return f"{count} {word}" + ("" if count == 1 else "s")


def _heic_note():
    """Warning when HEIC files are present but pillow-heif is missing."""
    heic = _count_files(sc.HEIC_EXTS)
    if heic and not _heif_available():
        return f" {heic} HEIC file(s) need pillow-heif."
    return ""


def _probe_images():
    count = _count_files(_IMAGE_EXTS)
    if count:
        return (f"Ready: {_plural(count, 'supported image')} in this folder."
                + _heic_note())
    return "Nothing to do: no supported images in this folder." + _heic_note()


def _probe_word_docs():
    count = _count_files(('.doc', '.docx'))
    if count:
        return f"Ready: {_plural(count, 'Word document')} in this folder."
    return "Nothing to do: no .doc/.docx files in this folder."


def _probe_pdfs():
    count = _count_files(sc.PDF_EXTS)
    if count:
        return f"Ready: {_plural(count, 'PDF')} in this folder."
    return "Nothing to do: no PDFs in this folder."


def _probe_merge_split():
    count = _count_files(sc.PDF_EXTS)
    if count >= 2:
        return f"Ready: {_plural(count, 'PDF')} - merge them or split one."
    if count == 1:
        return "Ready: 1 PDF - split it into pages (merging needs 2+)."
    return "Nothing to do: no PDFs in this folder."


def _probe_everything():
    counts = [
        ("Word", _count_files(sc.WORD_EXTS)),
        ("PDF", _count_files(sc.PDF_EXTS)),
        ("Excel", _count_files(sc.EXCEL_EXTS)),
        ("PowerPoint", _count_files(sc.POWERPOINT_EXTS)),
        ("image", _count_files(_IMAGE_EXTS)),
    ]
    total = sum(count for _label, count in counts)
    if total:
        summary = ", ".join(f"{count} {label}" for label, count in counts
                            if count)
        return f"Ready: {summary} in this folder." + _heic_note()
    return ("Nothing to combine yet: put Word, Excel, PowerPoint, PDF or "
            "image files in this folder.")


def _probe_loose_files():
    folder = os.getcwd()
    try:
        return sum(1 for f in os.listdir(folder)
                   if os.path.isfile(os.path.join(folder, f))
                   and not f.startswith("~$"))
    except OSError:
        return 0


def _probe_mixed_files():
    count = _probe_loose_files()
    if count:
        return f"Ready: {_plural(count, 'file')} in this folder."
    return "Nothing to organize: no loose files in this folder."


def _probe_phone_frame():
    folder = os.getcwd()
    frame_ok = os.path.isfile(os.path.join(folder, 'frame.png'))
    screens = os.path.join(folder, 'screens')
    if os.path.isdir(screens):
        count = sum(1 for f in os.listdir(screens)
                    if f.lower().endswith(_IMAGE_EXTS)
                    and not f.startswith('.'))
    else:
        count = 0
    frame_state = "frame.png found" if frame_ok else "frame.png will be created"
    if count:
        return f"Ready: {_plural(count, 'screenshot')} in screens/, {frame_state}."
    return "Nothing to do yet: put screenshots in screens/ " \
           f"({frame_state})."


@dataclass
class ScriptEntry:
    name: str
    category: str
    what: str
    needs: list
    how: list
    output: str
    probe: object
    func: object


SCRIPTS = [
    ScriptEntry(
        name="Images to PDF",
        category="PDF",
        what="Combine every image in the working folder into a single PDF.",
        needs=["Images (PNG, JPG, BMP, GIF, TIFF, WEBP, HEIC) in the "
               "working folder"],
        how=["Put your images in the working folder (D to change it)",
             "Select this script and press Enter or R"],
        output="output.pdf - one US Letter page per image, centered; "
               "photos are rotated per their EXIF orientation",
        probe=_probe_images,
        func=_SCRIPT_FUNCS["Images to PDF"],
    ),
    ScriptEntry(
        name="Images to PDF (Custom Aspect)",
        category="PDF",
        what="Combine images into one PDF where each page keeps the "
             "image's real size and aspect ratio.",
        needs=["Images (PNG, JPG, BMP, GIF, TIFF, WEBP, HEIC) in the "
               "working folder"],
        how=["Put your images in the working folder (D to change it)",
             "Press Enter or R and pick a page order (or press Esc for "
             "the default natural order)"],
        output="output.pdf - pages sized to each image (DPI-aware), "
               "ordered naturally (img2 before img10)",
        probe=_probe_images,
        func=_SCRIPT_FUNCS["Images to PDF (Custom Aspect)"],
    ),
    ScriptEntry(
        name="Word to PDF",
        category="PDF",
        what="Batch convert every Word document in the working folder to PDF.",
        needs=["Windows with Microsoft Word installed",
               ".doc/.docx files in the working folder"],
        how=["Put your Word documents in the working folder",
             "Select this script and press Enter or R"],
        output="A PDF next to each document; already-converted files are "
               "skipped, failed ones retried up to 3 times",
        probe=_probe_word_docs,
        func=_SCRIPT_FUNCS["Word to PDF"],
    ),
    ScriptEntry(
        name="PDF to Markdown",
        category="PDF",
        what="Convert every PDF in the working folder to a Markdown file.",
        needs=["Text-based PDFs in the working folder "
               "(scanned pages are detected and skipped)"],
        how=["Put your PDFs in the working folder",
             "Select this script and press Enter or R"],
        output="One .md file per PDF (headings, lists, tables, links); "
               "existing .md files are never overwritten",
        probe=_probe_pdfs,
        func=_SCRIPT_FUNCS["PDF to Markdown"],
    ),
    ScriptEntry(
        name="Image Format Converter",
        category="Image",
        what="Batch convert all images in the working folder to a format "
             "you pick: PNG, JPG, BMP, GIF, TIFF, WEBP or HEIC.",
        needs=["Images (PNG, JPG, BMP, GIF, TIFF, WEBP, HEIC) in the "
               "working folder"],
        how=["Put your images in the working folder",
             "Press Enter or R and pick the target format from the list"],
        output="A converted copy of each image in the working folder; "
               "originals are kept, files already in the target format "
               "are skipped",
        probe=_probe_images,
        func=_SCRIPT_FUNCS["Image Format Converter"],
    ),
    ScriptEntry(
        name="Image Resizer",
        category="Image",
        what="Batch resize or compress images: cap the longest side "
             "(1920/1280 px), scale to 50%, or squeeze JPG/WEBP copies "
             "under a target file size.",
        needs=["Images (PNG, JPG, BMP, GIF, TIFF, WEBP, HEIC) in the "
               "working folder"],
        how=["Put your images in the working folder",
             "Press Enter or R and pick a preset (Esc cancels)"],
        output="A resized or compressed copy of each image; originals are "
               "kept, target-size mode saves JPG copies",
        probe=_probe_images,
        func=_SCRIPT_FUNCS["Image Resizer"],
    ),
    ScriptEntry(
        name="Phone Frame Overlay",
        category="Image",
        what="Place your screenshots inside a phone frame image for "
             "polished mockups.",
        needs=["Screenshots inside a screens/ folder",
               "frame.png device mockup (both are created automatically "
               "on first run)"],
        how=["Select this script and press Enter or R once - it creates "
             "screens/ and a placeholder frame.png",
             "Drop your screenshots into screens/",
             "Run it again"],
        output="output/ folder with a framed copy of each screenshot "
               "(oversized screenshots are scaled to fit)",
        probe=_probe_phone_frame,
        func=_SCRIPT_FUNCS["Phone Frame Overlay"],
    ),
    ScriptEntry(
        name="File Organizer",
        category="Files",
        what="Sort the loose files in the working folder into subfolders "
             "by type (Images, Documents, ...) or by year-month.",
        needs=["Files in the working folder (folders are never touched)"],
        how=["Select this script, press Enter or R and pick By file type "
             "or By year-month"],
        output="Subfolders like Images/, Documents/ or 2025-01/ with the "
               "files moved inside; name collisions get a numbered suffix",
        probe=_probe_mixed_files,
        func=_SCRIPT_FUNCS["File Organizer"],
    ),
    ScriptEntry(
        name="Batch Renamer",
        category="Files",
        what="Rename all files in the working folder with one rule: "
             "sequence numbers, lowercase-with-dashes, or a YYYYMMDD "
             "date stamp.",
        needs=["Files in the working folder"],
        how=["Select this script, press Enter or R and pick a renaming "
             "rule",
             "Check the preview (old -> new) and confirm - nothing is "
             "changed until you do"],
        output="Renamed files in the working folder; name collisions get "
               "a numbered suffix",
        probe=_probe_mixed_files,
        func=_SCRIPT_FUNCS["Batch Renamer"],
    ),
]


# =============================================================================
# MODAL SCREENS
# =============================================================================

def _clamp_dialog_list(screen, list_id: str) -> None:
    """Cap a modal's list to what the terminal can actually show.

    The list must be the element that scrolls: if the dialog were clipped
    by the screen edge instead, the last options could be highlighted but
    invisible (there is no scrollbar on the dialog).
    """
    list_view = screen.query_one(f"#{list_id}", ListView)
    available = screen.size.height - 9  # margins, label, border, padding
    list_view.styles.max_height = max(4, min(20, available))


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
        _clamp_dialog_list(self, "select-modal-list")
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


class ThemePicker(ModalScreen):
    """Theme chooser with live preview: themes apply while you scroll.

    Dismisses with the chosen theme name, or None if cancelled
    (the app then reverts to the theme that was active before).
    """

    DEFAULT_THEME_NAME = DEFAULT_THEME

    def __init__(self, current: str) -> None:
        super().__init__()
        self._current = current
        self._preview = current
        self._names = []
        self._by_item = {}

    def compose(self) -> ComposeResult:
        themes = sorted(self.app.available_themes.values(),
                        key=lambda theme: theme.name)
        self._names = [theme.name for theme in themes]
        items = []
        for name in self._names:
            label = f"{name} (default)" if name == self.DEFAULT_THEME_NAME else name
            items.append(ListItem(Label(label)))
        yield Vertical(
            Label("Theme - scroll to preview, Enter to keep, Esc to cancel",
                  id="theme-picker-prompt"),
            ListView(*items, id="theme-picker-list"),
        )

    def on_mount(self) -> None:
        _clamp_dialog_list(self, "theme-picker-list")
        list_view = self.query_one("#theme-picker-list", ListView)
        for item, name in zip(list_view.children, self._names):
            self._by_item[item] = name
        list_view.focus()
        if self._current in self._names:
            list_view.index = self._names.index(self._current)

    @on(ListView.Highlighted)
    def preview(self, event: ListView.Highlighted) -> None:
        name = self._by_item.get(event.item)
        if name and name != self._preview:
            self._preview = name
            self.app.theme = name

    @on(ListView.Selected)
    def chosen(self, event: ListView.Selected) -> None:
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
    #script-filter {
        margin: 0 1 1 1;
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
    #script-detail {
        margin-top: 1;
        width: 1fr;
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
    SelectModal Vertical, ThemePicker Vertical {
        height: auto;
        margin: 2 8;
        padding: 1 1;
        background: $surface;
        border: round $accent;
    }
    #select-modal-prompt, #theme-picker-prompt {
        padding: 0 1;
    }
    #select-modal-list, #theme-picker-list {
        height: auto;
        max-height: 20;
    }
    """

    BINDINGS = [
        Binding("r", "run_highlighted", "Run"),
        Binding("f", "focus_filter", "Filter"),
        Binding("d", "change_directory", "Directory"),
        Binding("o", "open_folder", "Open folder"),
        Binding("t", "change_theme", "Theme"),
        Binding("c", "clear_log", "Clear log"),
        Binding("q", "quit", "Quit"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self._entries = {}
        self._saved_theme = DEFAULT_THEME

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True)
        with Horizontal(id="body"):
            with Vertical(id="sidebar"):
                yield Label("SCRIPTS", id="sidebar-header")
                yield Input(placeholder="Filter... (F)",
                            id="script-filter")
                items = [ListItem(Label(script.name)) for script in SCRIPTS]
                yield ListView(*items, id="script-list", initial_index=0)
            with Vertical(id="main"):
                yield Label("", id="script-title")
                yield Label("", id="script-category")
                yield Static("", id="script-detail", markup=True)
        yield RichLog(id="log-panel", markup=False, highlight=False, wrap=True)
        yield Footer()

    def get_system_commands(self, screen):
        for command in super().get_system_commands(screen):
            if command.title == "Screenshot":
                continue
            yield command

    def on_mount(self) -> None:
        settings = load_settings()
        theme = settings.get("theme")
        if theme and theme in self.available_themes:
            self.theme = theme
        last_dir = settings.get("last_dir")
        if last_dir and os.path.isdir(last_dir):
            try:
                os.chdir(last_dir)
            except OSError:
                pass
        list_view = self.query_one("#script-list", ListView)
        for item, script in zip(list_view.children, SCRIPTS):
            self._entries[item] = script
        self._refresh_directory()
        self._show_script(SCRIPTS[0])
        list_view.focus()
        self.log_line(f"Daily Tasks v{__version__} - ready.")
        self.log_markup("Select a script, press [bold]Enter[/] or [bold]R[/] to run, "
                        "[bold]F[/] to filter, [bold]D[/] to change the working "
                        "directory, [bold]T[/] for themes.")

    def _refresh_directory(self) -> None:
        self.sub_title = f"v{__version__} - {os.getcwd()}"

    def _show_script(self, script: ScriptEntry) -> None:
        self.query_one("#script-title", Label).update(script.name)
        self.query_one("#script-category", Label).update(f"[{script.category}]")
        self.query_one("#script-detail", Static).update(self._detail_markup(script))

    @staticmethod
    def _detail_markup(script: ScriptEntry) -> str:
        lines = ["[bold]WHAT[/]", f"  {script.what}", ""]
        lines.append("[bold]NEEDS[/]")
        lines.extend(f"  - {need}" for need in script.needs)
        lines.append("")
        lines.append("[bold]HOW[/]")
        lines.extend(f"  {number}. {step}"
                     for number, step in enumerate(script.how, 1))
        lines.append("")
        lines.append("[bold]OUTPUT[/]")
        lines.append(f"  {script.output}")
        lines.append("")
        status = script.probe()
        style = "green" if status.startswith("Ready") else "yellow"
        lines.append(f"[{style}]{status}[/]")
        return "\n".join(lines)

    def _refresh_detail(self) -> None:
        script = self._current_script()
        if script is not None:
            self._show_script(script)

    def _current_script(self):
        list_view = self.query_one("#script-list", ListView)
        if list_view.index is None or not list_view.children:
            return None
        item = list_view.children[list_view.index]
        return self._entries.get(item)

    # --- Filter -------------------------------------------------------------

    @on(Input.Changed, "#script-filter")
    def filter_changed(self, event: Input.Changed) -> None:
        self._apply_filter(event.value)

    @on(Input.Submitted, "#script-filter")
    def filter_submitted(self, event: Input.Submitted) -> None:
        list_view = self.query_one("#script-list", ListView)
        if list_view.children:
            list_view.focus()

    def _apply_filter(self, query: str) -> None:
        list_view = self.query_one("#script-list", ListView)
        list_view.clear()
        self._entries.clear()
        query = (query or "").strip().lower()
        for script in SCRIPTS:
            if query and query not in script.name.lower() \
                    and query not in script.category.lower():
                continue
            item = ListItem(Label(script.name))
            self._entries[item] = script
            list_view.append(item)
        if list_view.children:
            list_view.index = 0
        else:
            self.query_one("#script-title", Label).update("No match")
            self.query_one("#script-category", Label).update("")
            self.query_one("#script-detail", Static).update(
                "No scripts match your filter.")

    def action_focus_filter(self) -> None:
        self.query_one("#script-filter", Input).focus()

    def on_key(self, event) -> None:
        # Esc in the filter clears it and hands focus back to the list.
        if event.key == "escape":
            focused = self.focused
            if isinstance(focused, Input) and focused.id == "script-filter":
                if focused.value:
                    focused.value = ""
                list_view = self.query_one("#script-list", ListView)
                if list_view.children:
                    list_view.focus()

    # --- Log helpers ----------------------------------------------------------

    def log_line(self, line: str) -> None:
        self.query_one("#log-panel", RichLog).write(line)

    def log_markup(self, markup: str) -> None:
        self.query_one("#log-panel", RichLog).write(Text.from_markup(markup))

    # --- Selection / actions --------------------------------------------------

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
            self._refresh_detail()
            update_setting("last_dir", os.getcwd())
            self.log_line(f"Working directory changed to: {os.getcwd()}")
            self.notify(f"Working directory: {os.getcwd()}")
        else:
            self.log_line(f"Not a directory: {path}")
            self.notify(f"Not a directory: {path}", severity="error")

    def action_open_folder(self) -> None:
        folder = os.getcwd()
        if sys.platform != "win32":
            self.log_line(f"Opening folders is only supported on Windows: {folder}")
            return
        try:
            os.startfile(folder)  # noqa: S606
        except OSError as error:
            self.log_line(f"Could not open folder: {error}")

    def action_change_theme(self) -> None:
        self._saved_theme = self.theme
        self.push_screen(ThemePicker(self.theme), callback=self._theme_chosen)

    def _theme_chosen(self, choice) -> None:
        if choice is None:
            self.theme = self._saved_theme
            return
        update_setting("theme", choice)
        self.log_line(f"Theme set to: {choice}")

    # --- ask() bridge -----------------------------------------------------------

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

    # --- Script worker ----------------------------------------------------------

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
        self.app.call_from_thread(self._refresh_detail)


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
        # scripts/ ships as data files and is loaded at runtime
        "--add-data", f"scripts{os.pathsep}scripts",
        # Script files are data (not traced), so every library they import
        # must be requested explicitly. Prefer hidden-imports so PyInstaller's
        # own analysis (and library hooks) decide what else to include -
        # collect-all on pdfplumber/PIL would drag in optional extras
        # (pandas, torch, tkinter) installed on the build machine.
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--hidden-import", "PIL.ImageOps",
        "--hidden-import", "PIL.ImageDraw",
        "--hidden-import", "PIL.ExifTags",
        "--hidden-import", "pdfplumber",
        "--collect-submodules", "reportlab",
        "--collect-all", "pypdf",
        "--collect-all", "pillow_heif",
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

    # The dependency tree imports several heavy libraries only optionally
    # (guarded imports in rich/PIL/pdfplumber: IPython, pandas, numpy, ...).
    # PyInstaller follows them statically, so exclude everything the scripts
    # never use - without this the exe balloons to ~300 MB on machines with
    # torch/pandas installed.
    for excluded in (
        "torch", "torchvision", "torchaudio", "tensorflow", "keras",
        "onnxruntime", "safetensors", "tensorboard",
        "pandas", "scipy", "numpy", "numba", "llvmlite", "networkx", "sympy",
        "matplotlib", "PIL.ImageTk", "tkinter",
        "IPython", "jedi", "parso", "black", "nbformat", "nbconvert",
        "jsonschema", "zmq", "ipykernel", "jupyter_client", "jupyter_core",
        "notebook", "sqlalchemy", "fsspec", "dask", "bokeh", "panel",
        "holoviews", "pytest", "sphinx", "pyproj", "openpyxl",
    ):
        cmd.extend(["--exclude-module", excluded])

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
