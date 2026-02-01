"""
Images to PDF Converter (Custom Aspect Ratio)
Converts images to PDF preserving their native aspect ratio and DPI.
Supports multiple sorting options: natural, mtime, or EXIF date.
"""

import os
import re
from datetime import datetime
from PIL import Image, ImageOps, ExifTags
from reportlab.pdfgen import canvas


def natural_key(s: str):
    """Sort key for natural sorting (e.g., Q2 before Q10)."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]


def exif_datetime(path: str):
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


def images_to_pdf(folder_path: str, output_pdf: str, default_dpi: int = 72, 
                  order_by: str = "natural", reverse: bool = False) -> None:
    """
    Convert all images in folder_path to a PDF with native aspect ratios.
    
    Args:
        folder_path: Directory containing images
        output_pdf: Output PDF file path
        default_dpi: Default DPI if not in image metadata
        order_by: Sort method - "natural", "mtime", or "exif"
        reverse: Reverse sort order
    """
    extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff')
    images = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
              if f.lower().endswith(extensions)]

    if not images:
        print("No images found in the folder.")
        return

    # Sort images
    if order_by == "natural":
        images.sort(key=lambda p: natural_key(os.path.basename(p)), reverse=reverse)
    elif order_by == "mtime":
        images.sort(key=lambda p: os.path.getmtime(p), reverse=reverse)
    elif order_by == "exif":
        def exif_or_mtime(p):
            dt = exif_datetime(p)
            return dt.timestamp() if dt else os.path.getmtime(p)
        images.sort(key=exif_or_mtime, reverse=reverse)
    else:
        images.sort(key=lambda p: os.path.basename(p).lower(), reverse=reverse)

    c = canvas.Canvas(output_pdf)

    for img_path in images:
        print(f"Processing: {os.path.basename(img_path)}")
        
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
    print(f"PDF saved: {output_pdf}")


def main():
    """Run the converter on the current directory."""
    folder_path = os.path.dirname(os.path.realpath(__file__))
    output_pdf = os.path.join(folder_path, "output.pdf")
    
    print(f"Folder: {folder_path}")
    print(f"Output: {output_pdf}")
    
    # Options: "natural" (default), "mtime", or "exif"
    images_to_pdf(folder_path, output_pdf, order_by="natural")
    print("Done!")


if __name__ == "__main__":
    main()
