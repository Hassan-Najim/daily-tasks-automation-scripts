"""
Everything to PDF
Combines every convertible file in the working folder into a single PDF:
Word (.doc/.docx), Excel (.xls/.xlsx), PowerPoint (.ppt/.pptx), existing
PDFs, and images of any supported type - merged in name order into
combined.pdf.

Office files are exported with Microsoft Office via COM (Windows only);
files whose Office app is not installed are skipped with a warning.
Images become pages at their native size (DPI-aware) with EXIF rotation.
"""

import os
import shutil
import sys
import tempfile

try:
    import script_common as sc
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
    import script_common as sc

OUTPUT_NAME = "combined.pdf"


def _office_app_available(prog_id):
    """True if the Office app (e.g. "Word.Application") can be started."""
    if sys.platform != "win32":
        return False
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx(prog_id)
        app.Quit()
        return True
    except Exception:
        return False
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _image_to_pdf(img_path, pdf_path, default_dpi=72):
    """Render one image as a single-page PDF at its native (DPI-aware) size."""
    from PIL import Image, ImageOps
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    with Image.open(img_path) as im:
        im = ImageOps.exif_transpose(im)
        im.load()
        px_w, px_h = im.size
        dpi_x, dpi_y = im.info.get("dpi", (default_dpi, default_dpi))
        page_w = px_w * 72.0 / (dpi_x or default_dpi)
        page_h = px_h * 72.0 / (dpi_y or default_dpi)

        c = canvas.Canvas(pdf_path, pagesize=(page_w, page_h))
        try:
            c.drawImage(ImageReader(im), 0, 0, width=page_w, height=page_h,
                        mask='auto')
            c.showPage()
        finally:
            c.save()


# prog_id per Office file kind, with a human label
_OFFICE_KINDS = (
    (sc.WORD_EXTS, "Word.Application", "Word"),
    (sc.EXCEL_EXTS, "Excel.Application", "Excel"),
    (sc.POWERPOINT_EXTS, "PowerPoint.Application", "PowerPoint"),
)


def main(log=print, ask=None):
    folder = os.getcwd()
    heic = sc.register_heif()
    image_exts = sc.IMAGE_EXTS + (sc.HEIC_EXTS if heic else ())

    # Collect convertible sources in name order
    sources = []  # (filename, kind)
    for name in sorted(os.listdir(folder), key=sc.natural_key):
        lower = name.lower()
        if name.startswith("~$") or lower == OUTPUT_NAME:
            continue
        path = os.path.join(folder, name)
        if not os.path.isfile(path):
            continue
        if lower.endswith(image_exts):
            sources.append((name, "image"))
        elif lower.endswith(sc.PDF_EXTS):
            sources.append((name, "pdf"))
        else:
            for exts, _prog_id, label in _OFFICE_KINDS:
                if lower.endswith(exts):
                    sources.append((name, label))
                    break

    if not sources:
        log("Nothing to combine: no Word, Excel, PowerPoint, PDF or image "
            "files in this folder.")
        return

    log(f"Folder: {folder}")
    log(f"Will combine {len(sources)} file(s) into {OUTPUT_NAME}:")
    for name, kind in sources:
        log(f"  - {name} ({kind})")
    log("")

    temp_dir = tempfile.mkdtemp(prefix="everything_pdf_")
    output_path = os.path.join(folder, OUTPUT_NAME)
    merged = []          # ordered list of PDFs to merge
    failed = []
    unavailable = {}     # label -> reason

    try:
        for name, kind in sources:
            src_path = os.path.join(folder, name)
            try:
                if kind == "image":
                    temp_pdf = os.path.join(
                        temp_dir, f"{len(merged):04d}_img.pdf")
                    _image_to_pdf(src_path, temp_pdf)
                    merged.append(temp_pdf)
                    log(f"Added image: {name}")
                elif kind == "pdf":
                    merged.append(src_path)
                    log(f"Added PDF: {name}")
                else:
                    if kind in unavailable:
                        log(f"Skipped ({kind} unavailable): {name}")
                        failed.append((name, f"{kind} not available"))
                        continue
                    temp_pdf = os.path.join(
                        temp_dir, f"{len(merged):04d}_office.pdf")
                    prog_id = next(prog for exts, prog, label
                                   in _OFFICE_KINDS if label == kind)
                    try:
                        sc.office_export_pdf(src_path, temp_pdf, prog_id,
                                             log=log)
                    except Exception as e:
                        # First failure per app: is the app itself missing?
                        if kind not in unavailable and \
                                not _office_app_available(prog_id):
                            unavailable[kind] = True
                            log(f"WARNING: Microsoft {kind} is not "
                                f"installed - skipping {kind} files.")
                            log(f"Skipped: {name}")
                            failed.append((name, f"{kind} not installed"))
                        else:
                            log(f"Failed to convert {name}: {e}")
                            failed.append((name, str(e)))
                        continue
                    merged.append(temp_pdf)
                    log(f"Converted with {kind}: {name}")
            except Exception as e:
                log(f"Failed: {name}: {e}")
                failed.append((name, str(e)))

        if not merged:
            log("\nNothing could be converted - no PDF created.")
            return

        # Merge everything in source order
        from pypdf import PdfWriter
        writer = PdfWriter()
        for pdf_path in merged:
            try:
                writer.append(pdf_path)
            except Exception as e:
                log(f"Warning: could not merge a converted file: {e}")
        with open(output_path, "wb") as f:
            writer.write(f)

        log(f"\nSaved: {output_path} ({len(writer.pages)} pages from "
            f"{len(merged)} file(s))")
        if failed:
            log(f"Skipped/failed: {len(failed)}")
            for name, err in failed:
                log(f"  - {name}: {err}")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
