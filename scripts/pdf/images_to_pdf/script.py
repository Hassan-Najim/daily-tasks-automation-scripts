"""
Images to PDF Converter
Converts all images in the working folder into a single PDF file, one
US Letter page per image, centered. Photos are rotated according to their
EXIF orientation so phone pictures come out upright.
"""

import os
import sys

try:
    import script_common as sc
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
    import script_common as sc

from PIL import Image, ImageOps
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

OUTPUT_NAME = "output.pdf"


def images_to_pdf(folder_path, output_pdf, log=print):
    """Convert all supported images in folder_path into one PDF."""
    heic = sc.register_heif()
    extensions = sc.IMAGE_EXTS + (sc.HEIC_EXTS if heic else ())
    images = sorted(
        (f for f in os.listdir(folder_path)
         if f.lower().endswith(extensions) and not f.startswith("~$")),
        key=sc.natural_key)

    if not images:
        log("No images found in this folder.")
        return 0

    width, height = letter
    c = canvas.Canvas(output_pdf, pagesize=letter)
    converted = 0
    try:
        for filename in images:
            img_path = os.path.join(folder_path, filename)
            log(f"Processing: {filename}")
            try:
                with Image.open(img_path) as image:
                    image = ImageOps.exif_transpose(image)
                    image.load()
                    aspect = image.width / float(image.height)
            except Exception as e:
                log(f"  Failed to read {filename}: {e}")
                continue

            if aspect > 1:  # landscape: fit page width
                img_width = width
                img_height = width / aspect
            else:  # portrait: fit page height
                img_height = height
                img_width = height * aspect

            x = (width - img_width) / 2
            y = (height - img_height) / 2

            # Draw from the (already EXIF-corrected) image object, not the
            # raw file, so reportlab honors the rotation.
            c.drawImage(ImageReader(image), x, y, img_width, img_height,
                        mask='auto')
            c.showPage()
            converted += 1
    finally:
        c.save()

    return converted


def main(log=print, ask=None):
    folder_path = os.getcwd()
    output_pdf = os.path.join(folder_path, OUTPUT_NAME)

    log(f"Folder: {folder_path}")
    log(f"Output: {output_pdf}")
    converted = images_to_pdf(folder_path, output_pdf, log=log)
    if converted:
        log(f"\nPDF saved: {output_pdf} ({converted} page(s))")
        log("Note: animated GIFs use their first frame.")


if __name__ == "__main__":
    main()
