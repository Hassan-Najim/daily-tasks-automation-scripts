"""
PDF Merge & Split
Merge every PDF in the working folder into one file (merged.pdf, name
order), or split one PDF into single-page PDFs inside a split/ folder.
"""

import os
import sys

try:
    import script_common as sc
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
    import script_common as sc

OUTPUT_NAME = "merged.pdf"
SPLIT_DIR = "split"

_MERGE = "Merge all PDFs into one"
_SPLIT = "Split one PDF into single pages"


def _pdf_list(folder):
    return sorted(
        (f for f in os.listdir(folder)
         if f.lower().endswith(".pdf") and not f.startswith("~$")),
        key=sc.natural_key)


def _merge(folder, pdfs, log=print):
    from pypdf import PdfWriter

    output_path = os.path.join(folder, OUTPUT_NAME)
    writer = PdfWriter()
    for name in pdfs:
        path = os.path.join(folder, name)
        try:
            before = len(writer.pages)
            writer.append(path)
            log(f"Merged: {name} ({len(writer.pages) - before} pages)")
        except Exception as e:
            log(f"Failed to read {name}: {e}")
    if len(writer.pages) == 0:
        log("\nNo pages could be read - no file written.")
        return
    with open(output_path, "wb") as f:
        writer.write(f)
    log(f"\nSaved: {output_path} ({len(writer.pages)} pages from "
        f"{len(pdfs)} PDF(s))")


def _split(folder, pdfs, ask, log=print):
    from pypdf import PdfReader, PdfWriter

    name = pdfs[0]
    if len(pdfs) > 1:
        if ask is not None:
            name = ask("Which PDF should be split?", pdfs)
        else:
            log("Which PDF should be split?")
            for i, option in enumerate(pdfs, 1):
                log(f"  {i}. {option}")
            raw = input("Enter number: ").strip()
            if not (raw.isdigit() and 1 <= int(raw) <= len(pdfs)):
                log("Cancelled.")
                return
            name = pdfs[int(raw) - 1]
        if name is None:
            log("Cancelled.")
            return

    source = os.path.join(folder, name)
    try:
        reader = PdfReader(source)
    except Exception as e:
        log(f"Could not open {name}: {e}")
        return

    output_dir = os.path.join(folder, SPLIT_DIR)
    os.makedirs(output_dir, exist_ok=True)

    base = os.path.splitext(name)[0]
    width = max(3, len(str(len(reader.pages))))
    written = 0
    for index, page in enumerate(reader.pages, 1):
        out_name = f"{base}_page_{index:0{width}d}.pdf"
        writer = PdfWriter()
        writer.add_page(page)
        with open(os.path.join(output_dir, out_name), "wb") as f:
            writer.write(f)
        written += 1
    log(f"\nSaved {written} page(s) into {SPLIT_DIR}/ "
        f"({base}_page_001.pdf, ...)")


def main(log=print, ask=None):
    folder = os.getcwd()
    pdfs = [f for f in _pdf_list(folder) if f.lower() != OUTPUT_NAME]

    if not pdfs:
        log("No PDF files found in this folder.")
        return

    mode = None
    if ask is not None:
        mode = ask("What do you want to do?", [_MERGE, _SPLIT])
    else:
        log("What do you want to do?")
        log(f"  1. {_MERGE}")
        log(f"  2. {_SPLIT}")
        raw = input("Choice [1-2]: ").strip()
        if raw == "1":
            mode = _MERGE
        elif raw == "2":
            mode = _SPLIT
    if not mode:
        log("Cancelled.")
        return

    if mode == _MERGE:
        if len(pdfs) < 2:
            log("Merging needs at least 2 PDFs (use Split for one).")
            return
        log(f"Merging {len(pdfs)} PDF(s) in name order...")
        _merge(folder, pdfs, log=log)
    else:
        _split(folder, pdfs, ask, log=log)


if __name__ == "__main__":
    main()
