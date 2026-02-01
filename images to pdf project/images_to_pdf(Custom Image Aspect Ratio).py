import os
import re
from datetime import datetime
from PIL import Image, ImageOps, ExifTags
from reportlab.pdfgen import canvas

def natural_key(s: str):
    # Split into text/number chunks: "Q10.png" -> ["q", 10, ".png"]
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]

def exif_datetime(path: str):
    # Return EXIF DateTimeOriginal if present; else None
    try:
        with Image.open(path) as im:
            exif = im.getexif()
            if not exif:
                return None
            # Find the tag id for DateTimeOriginal
            dt_tag = None
            for k, v in ExifTags.TAGS.items():
                if v == "DateTimeOriginal":
                    dt_tag = k
                    break
            if dt_tag is None or dt_tag not in exif:
                return None
            dt_str = exif.get(dt_tag)  # format "YYYY:MM:DD HH:MM:SS"
            if isinstance(dt_str, bytes):
                dt_str = dt_str.decode(errors="ignore")
            return datetime.strptime(dt_str, "%Y:%m:%d %H:%M:%S")
    except Exception:
        return None

def images_to_pdf(folder_path, output_pdf, default_dpi=72, order_by="natural", reverse=False):
    # Collect full paths so sort keys can use path easily
    exts = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff')
    images = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
              if f.lower().endswith(exts)]

    if not images:
        print("No images found in the folder.")
        return

    # ------- Deterministic ordering -------
    if order_by == "natural":
        images.sort(key=lambda p: natural_key(os.path.basename(p)), reverse=reverse)
    elif order_by == "mtime":
        images.sort(key=lambda p: os.path.getmtime(p), reverse=reverse)
    elif order_by == "exif":
        # Sort by EXIF DateTimeOriginal; fallback to mtime if missing
        def exif_or_mtime(p):
            dt = exif_datetime(p)
            return dt.timestamp() if dt else os.path.getmtime(p)
        images.sort(key=exif_or_mtime, reverse=reverse)
    else:
        # Fallback: plain lexicographic
        images.sort(key=lambda p: os.path.basename(p).lower(), reverse=reverse)
    # -------------------------------------

    c = canvas.Canvas(output_pdf)

    for img_path in images:
        print(f"Processing image: {img_path}")
        with Image.open(img_path) as im:
            im = ImageOps.exif_transpose(im)

            px_w, px_h = im.size
            dpi_x, dpi_y = im.info.get("dpi", (default_dpi, default_dpi))  # tuple or missing

            page_w = px_w * 72.0 / (dpi_x or default_dpi)
            page_h = px_h * 72.0 / (dpi_y or default_dpi)

            c.setPageSize((page_w, page_h))
            c.drawImage(img_path, 0, 0, width=page_w, height=page_h,
                        preserveAspectRatio=False, mask='auto')
            c.showPage()

    c.save()
    print(f"PDF saved as {output_pdf}")

if __name__ == "__main__":
    folder_path = os.path.dirname(os.path.realpath(__file__))
    print(f"Folder path: {folder_path}")
    output_pdf = os.path.join(folder_path, "output.pdf")
    print(f"Output PDF path: {output_pdf}")

    # Choose one: "natural" (default), "mtime", or "exif"
    images_to_pdf(folder_path, output_pdf, order_by="natural")
    print("Finished processing")
