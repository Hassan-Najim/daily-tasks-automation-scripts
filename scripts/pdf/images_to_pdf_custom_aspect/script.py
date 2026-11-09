"""
Images to PDF Converter (Custom Aspect Ratio)
Converts images to a single PDF where each page keeps the image's native
aspect ratio and DPI (DPI-aware page sizing). Photos are rotated according
to their EXIF orientation. Supports ordering by name, modified date, or
EXIF date taken.
"""

import os
import sys
from datetime import datetime

try:
    import script_common as sc
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
    import script_common as sc

from PIL import Image, ImageOps, ExifTags
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

OUTPUT_NAME = "output.pdf"

_ORDER_CHOICES = [
    ("Natural (1, 2, 10)", "natural"),
    ("Name (A-Z)", "name"),
    ("Date modified", "mtime"),
    ("EXIF date taken", "exif"),
]


def exif_datetime(path):
    """Return EXIF DateTimeOriginal if present, else None."""
    try:
        with Image.open(path) as im:
            exif = im.getexif()
            if not exif:
                return None
            dt_tag = None
            for k, v in ExifTags.TAGS.items():
                if v == "DateTimeOriginal":
                    dt_tag = k
                    break
            if dt_tag is None or dt_tag not in exif:
                return None
            dt_str = exif.get(dt_tag)
            if isinstance(dt_str, bytes):
                dt_str = dt_str.decode(errors="ignore")
            return datetime.strptime(dt_str, "%Y:%m:%d %H:%M:%S")
    except Exception:
        return None


def _sort_images(images, order_by, reverse):
    """Sort image paths by the requested order."""
    if order_by == "natural":
        images.sort(key=lambda p: sc.natural_key(os.path.basename(p)),
                    reverse=reverse)
    elif order_by == "mtime":
        images.sort(key=os.path.getmtime, reverse=reverse)
    elif order_by == "exif":
        def exif_or_mtime(p):
            dt = exif_datetime(p)
            return dt.timestamp() if dt else os.path.getmtime(p)
        images.sort(key=exif_or_mtime, reverse=reverse)
    else:  # plain name
        images.sort(key=lambda p: os.path.basename(p).lower(), reverse=reverse)


def images_to_pdf(folder_path, output_pdf, default_dpi=72,
                  order_by="natural", reverse=False, log=print):
    """Convert all images in folder_path to a PDF with native page sizes."""
    heic = sc.register_heif()
    extensions = sc.IMAGE_EXTS + (sc.HEIC_EXTS if heic else ())
    images = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
              if f.lower().endswith(extensions) and not f.startswith("~$")]

    if not images:
        log("No images found in this folder.")
        return 0

    _sort_images(images, order_by, reverse)
    c = canvas.Canvas(output_pdf)
    converted = 0
    try:
        for img_path in images:
            log(f"Processing: {os.path.basename(img_path)}")
            try:
                with Image.open(img_path) as im:
                    im = ImageOps.exif_transpose(im)
                    im.load()
                    px_w, px_h = im.size
                    dpi_x, dpi_y = im.info.get("dpi",
                                               (default_dpi, default_dpi))

                    page_w = px_w * 72.0 / (dpi_x or default_dpi)
                    page_h = px_h * 72.0 / (dpi_y or default_dpi)

                    c.setPageSize((page_w, page_h))
                    c.drawImage(ImageReader(im), 0, 0, width=page_w,
                                height=page_h, preserveAspectRatio=False,
                                mask='auto')
                    c.showPage()
                    converted += 1
            except Exception as e:
                log(f"  Failed: {e}")
    finally:
        c.save()

    return converted


def main(log=print, ask=None):
    folder_path = os.getcwd()
    output_pdf = os.path.join(folder_path, OUTPUT_NAME)

    order_by = "natural"
    if ask is not None:
        label = ask("Page order", [label for label, _ in _ORDER_CHOICES])
        if label is None:
            log("Cancelled.")
            return
        order_by = dict(_ORDER_CHOICES)[label]
    else:
        log("Order: natural (rename order_by to change)")

    log(f"Folder: {folder_path}")
    log(f"Output: {output_pdf}")
    converted = images_to_pdf(folder_path, output_pdf, order_by=order_by,
                              log=log)
    if converted:
        log(f"\nPDF saved: {output_pdf} ({converted} page(s))")


if __name__ == "__main__":
    main()
