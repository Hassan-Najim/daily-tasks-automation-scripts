"""
Batch Renamer
Renames all files in the working folder with one rule: sequence numbers
("name 01.ext"), lowercase-with-dashes, or a YYYYMMDD date-stamp prefix.
Every change is previewed first - nothing is renamed until you confirm.
"""

import os
import re
import sys
from datetime import datetime

try:
    import script_common as sc
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
    import script_common as sc

NUMBER = "Number the files (name 01.ext, name 02.ext, ...)"
LOWDASH = "Lowercase with dashes (My File (2).JPG -> my-file-2.jpg)"
DATESTAMP = "Date-stamp names (name.ext -> 20251109_name.ext)"


def _plan_number(files):
    """Append ' NN' before the extension by folder position (global order)."""
    width = max(2, len(str(len(files))))
    plan = []
    for position, name in enumerate(files, 1):
        base, ext = os.path.splitext(name)
        plan.append((name, f"{base} {position:0{width}d}{ext}"))
    return plan


def _plan_lowdash(files):
    plan = []
    for name in files:
        base, ext = os.path.splitext(name)
        new_base = re.sub(r"\s+", "-", base.strip().lower())
        new_base = re.sub(r"-+", "-", new_base).strip("-")
        new_base = re.sub(r"[^a-z0-9\-_.]", "", new_base) or "file"
        plan.append((name, new_base + ext.lower()))
    return plan


def _plan_datestamp(files):
    plan = []
    for name in files:
        base, ext = os.path.splitext(name)
        mtime = os.path.getmtime(os.path.join(os.getcwd(), name))
        stamp = datetime.fromtimestamp(mtime).strftime("%Y%m%d")
        plan.append((name, f"{stamp}_{base}{ext}"))
    return plan


_PLANS = {NUMBER: _plan_number, LOWDASH: _plan_lowdash,
          DATESTAMP: _plan_datestamp}


def _resolve_collisions(folder, plan):
    """Ensure no new name collides with another new name or existing file."""
    taken = set()
    resolved = []
    for old, new in plan:
        candidate = new
        base, ext = os.path.splitext(new)
        n = 1
        while (candidate.lower() in taken
               or (os.path.exists(os.path.join(folder, candidate))
                   and candidate.lower() != old.lower())):
            candidate = f"{base}_{n}{ext}"
            n += 1
        taken.add(candidate.lower())
        resolved.append((old, candidate))
    return resolved


def main(log=print, ask=None):
    folder = os.getcwd()

    files = sorted(
        (f for f in os.listdir(folder)
         if os.path.isfile(os.path.join(folder, f))
         and not f.startswith("~$") and not f.startswith(".")),
        key=sc.natural_key)
    if not files:
        log("No files in this folder to rename.")
        return

    modes = list(_PLANS)
    mode = None
    if ask is not None:
        mode = ask("Pick a renaming rule", modes)
    else:
        log("Pick a renaming rule:")
        for i, label in enumerate(modes, 1):
            log(f"  {i}. {label}")
        raw = input(f"Choice [1-{len(modes)}]: ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(modes):
            mode = modes[int(raw) - 1]
    if not mode:
        log("Cancelled.")
        return

    plan = _resolve_collisions(folder, _PLANS[mode](files))

    log(f"\nPreview - {len(plan)} file(s) would be renamed:")
    changes = 0
    for old, new in plan:
        if old == new:
            log(f"  (unchanged) {old}")
            continue
        log(f"  {old}  ->  {new}")
        changes += 1

    if changes == 0:
        log("\nNothing to rename - all names already match the rule.")
        return

    if ask is not None:
        confirmed = ask(f"Apply {changes} rename(s)?",
                        ["Apply", "Cancel"]) == "Apply"
    else:
        log(f"\nApply {changes} rename(s)? (y/n)")
        confirmed = input("> ").strip().lower() in ("y", "yes")
    if not confirmed:
        log("Cancelled - nothing was renamed.")
        return

    renamed = 0
    for old, new in plan:
        if old == new:
            continue
        try:
            os.rename(os.path.join(folder, old), os.path.join(folder, new))
            renamed += 1
        except OSError as e:
            log(f"Failed: {old}: {e}")
    log(f"\nRenamed {renamed} file(s).")


if __name__ == "__main__":
    main()
