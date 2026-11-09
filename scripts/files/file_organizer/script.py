"""
File Organizer
Sorts the loose files in the working folder into subfolders, either by
file type (Images, Documents, Archives, ...) or by year-month from the
date modified. Folders themselves are never touched.
"""

import os
import shutil
import sys
from datetime import datetime

try:
    import script_common as sc
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
    import script_common as sc

BY_TYPE = "Sort by file type"
BY_DATE = "Sort by year-month (date modified)"

# Folder name -> extensions
TYPE_MAP = {
    "Images": sc.IMAGE_EXTS + sc.HEIC_EXTS,
    "Documents": ('.pdf', '.docx', '.doc', '.txt', '.md', '.rtf', '.odt',
                  '.pages'),
    "Spreadsheets": ('.xls', '.xlsx', '.csv', '.ods', '.numbers'),
    "Presentations": ('.ppt', '.pptx', '.odp', '.key'),
    "Archives": ('.zip', '.rar', '.7z', '.tar', '.gz', '.bz2'),
    "Audio": ('.mp3', '.wav', '.flac', '.m4a', '.ogg', '.wma'),
    "Video": ('.mp4', '.mov', '.avi', '.mkv', '.webm', '.wmv'),
    "Code": ('.py', '.js', '.ts', '.html', '.css', '.java', '.c', '.cpp',
             '.cs', '.json', '.xml', '.sh', '.ps1', '.bat', '.sql'),
    "Installers": ('.exe', '.msi', '.apk'),
}


def _collision_free(directory, filename):
    """Return a path in directory that does not exist yet (adds _1, _2...)."""
    candidate = os.path.join(directory, filename)
    if not os.path.exists(candidate):
        return candidate
    base, ext = os.path.splitext(filename)
    n = 1
    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{base}_{n}{ext}")
        n += 1
    return candidate


def _folder_for(filename):
    lower = filename.lower()
    for folder, exts in TYPE_MAP.items():
        if lower.endswith(exts):
            return folder
    return "Other"


def _organize_by_type(folder, files, log=print):
    moved = 0
    for name in files:
        dest_dir = os.path.join(folder, _folder_for(name))
        os.makedirs(dest_dir, exist_ok=True)
        dest = _collision_free(dest_dir, name)
        shutil.move(os.path.join(folder, name), dest)
        log(f"{name} -> {_folder_for(name)}/{os.path.basename(dest)}")
        moved += 1
    return moved


def _organize_by_date(folder, files, log=print):
    moved = 0
    for name in files:
        mtime = os.path.getmtime(os.path.join(folder, name))
        month = datetime.fromtimestamp(mtime).strftime("%Y-%m")
        dest_dir = os.path.join(folder, month)
        os.makedirs(dest_dir, exist_ok=True)
        dest = _collision_free(dest_dir, name)
        shutil.move(os.path.join(folder, name), dest)
        log(f"{name} -> {month}/{os.path.basename(dest)}")
        moved += 1
    return moved


def main(log=print, ask=None):
    folder = os.getcwd()

    files = sorted(
        (f for f in os.listdir(folder)
         if os.path.isfile(os.path.join(folder, f))
         and not f.startswith("~$") and not f.startswith(".")),
        key=sc.natural_key)
    if not files:
        log("No loose files in this folder - nothing to organize.")
        return

    mode = None
    if ask is not None:
        mode = ask("How should the files be sorted?", [BY_TYPE, BY_DATE])
    else:
        log("How should the files be sorted?")
        log(f"  1. {BY_TYPE}")
        log(f"  2. {BY_DATE}")
        raw = input("Choice [1-2]: ").strip()
        if raw == "1":
            mode = BY_TYPE
        elif raw == "2":
            mode = BY_DATE
    if not mode:
        log("Cancelled.")
        return

    log(f"Folder: {folder}")
    log(f"Files: {len(files)}\n")
    moved = (_organize_by_type(folder, files, log=log) if mode == BY_TYPE
             else _organize_by_date(folder, files, log=log))
    log(f"\nMoved {moved} file(s) into subfolders.")


if __name__ == "__main__":
    main()
