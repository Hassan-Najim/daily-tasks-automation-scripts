"""
Word to PDF Converter
Batch converts every Word document (.doc/.docx) in the working folder to a
PDF next to it, using Microsoft Word. Already-converted files are skipped.
"""

import os
import sys

try:
    import script_common as sc
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
    import script_common as sc


def main(log=print, ask=None):
    if sys.platform != "win32":
        log("Error: this script needs Windows with Microsoft Word installed.")
        return

    folder = os.getcwd()
    files = sorted(
        (f for f in os.listdir(folder)
         if f.lower().endswith((".docx", ".doc")) and not f.startswith("~$")),
        key=sc.natural_key)

    if not files:
        log("No .doc or .docx files found in this folder.")
        return

    failures = []
    converted = 0
    for name in files:
        pdf_path = os.path.join(folder, os.path.splitext(name)[0] + ".pdf")
        if os.path.exists(pdf_path):
            log(f"Skipping (exists): {name}")
            continue
        log(f"Converting: {name}")
        try:
            sc.office_export_pdf(os.path.join(folder, name), pdf_path,
                                 "Word.Application", log=log)
            converted += 1
        except Exception as e:
            log(f"  FAILED: {name}")
            failures.append((name, str(e)))

    log(f"\nConverted: {converted}")
    if failures:
        log(f"Failed: {len(failures)}")
        for name, err in failures:
            log(f"  - {name}: {err}")
    else:
        log("No failures.")


if __name__ == "__main__":
    main()
