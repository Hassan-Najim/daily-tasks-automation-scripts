"""
Image Resizer
Batch resize or compress images in the working folder: cap the longest
side (1920/1280 px), scale to 50%, or create JPG/WEBP copies under a
target file size. Originals are always kept.
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

# (label, kind, value)
_CHOICES = [
    ("Longest side -> 1920 px", "max_side", 1920),
    ("Longest side -> 1280 px", "max_side", 1280),
    ("Scale to 50 %", "scale", 50),
    ("JPG copies under 1 MB", "target", 1024 * 1024),
    ("JPG copies under 500 KB", "target", 500 * 1024),
]


def _flatten_for_jpg(img):
    """Convert to RGB for JPEG saving, compositing alpha onto white."""
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        background = Image.new("RGB", img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[-1])
        return background
    if img.mode != "RGB":
        return img.convert("RGB")
    return img


def _save_mode(img, target_format):
    """Adjust the image mode for the target format (converter rules)."""
    if target_format in ("JPG", "JPEG"):
        return _flatten_for_jpg(img)
    if target_format == "GIF":
        return img.convert("P") if img.mode not in ("P", "L", "RGB") else img
    if img.mode in ("RGBA", "LA", "P"):
        return img.convert("RGBA")
    if img.mode not in ("RGB", "L"):
        return img.convert("RGB")
    return img


def _ext_for(save_format):
    return {"JPEG": ".jpg", "PNG": ".png", "WEBP": ".webp"}.get(
        save_format, ".png")


def _compress_under_target(img, out_path, target_bytes):
    """Binary-search the JPEG quality that fits under target_bytes.

    Progressively shrinks the image if even the lowest quality is too big.
    Returns (quality or None, final_size).
    """
    scale = 1.0
    for _attempt in range(5):
        candidate = img
        if scale < 1.0:
            size = (max(1, int(img.width * scale)),
                    max(1, int(img.height * scale)))
            candidate = img.resize(size, Image.LANCZOS)
        lo, hi = 1, 95
        while lo <= hi:
            quality = (lo + hi) // 2
            candidate.save(out_path, "JPEG", quality=quality, optimize=True)
            if os.path.getsize(out_path) <= target_bytes:
                lo = quality + 1
            else:
                hi = quality - 1
        best_quality = hi
        if best_quality >= 1:
            candidate.save(out_path, "JPEG", quality=best_quality,
                           optimize=True)
            return best_quality, os.path.getsize(out_path)
        scale *= 0.7  # still too big even at quality 1 - shrink and retry
    return None, os.path.getsize(out_path)


def main(log=print, ask=None):
    folder = os.getcwd()
    sc.register_heif()

    images = sorted(
        (f for f in os.listdir(folder)
         if f.lower().endswith(sc.IMAGE_EXTS + sc.HEIC_EXTS)
         and not f.startswith("~$")),
        key=sc.natural_key)
    if not images:
        log("No supported images found in this folder.")
        return

    labels = [label for label, _kind, _value in _CHOICES]
    choice = None
    if ask is not None:
        choice = ask("Pick a preset", labels)
    else:
        log("Pick a preset:")
        for i, label in enumerate(labels, 1):
            log(f"  {i}. {label}")
        raw = input(f"Choice [1-{len(labels)}]: ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(labels):
            choice = labels[int(raw) - 1]
    if not choice:
        log("Cancelled.")
        return

    _kind, value = next((kind, val) for label, kind, val in _CHOICES
                        if label == choice)
    log(f"Preset: {choice}")
    log(f"Images: {len(images)} (originals are kept)\n")

    done = failed = skipped = 0
    for filename in images:
        src = os.path.join(folder, filename)
        base, ext = os.path.splitext(filename)
        try:
            with Image.open(src) as img:
                img = ImageOps.exif_transpose(img)
                img.load()
                original_size = os.path.getsize(src)

                if _kind == "target":
                    out_path = os.path.join(folder, f"{base}_resized.jpg")
                    if os.path.abspath(out_path) == os.path.abspath(src):
                        skipped += 1
                        continue
                    quality, size = _compress_under_target(
                        _flatten_for_jpg(img), out_path, value)
                    if quality is None:
                        log(f"Could not fit {filename} under target "
                            f"({size / 1024:.0f} KB at minimum quality)")
                    else:
                        log(f"Compressed: {filename} -> "
                            f"{os.path.basename(out_path)} "
                            f"({original_size // 1024} KB -> "
                            f"{size // 1024} KB, quality {quality})")
                    done += 1
                    continue

                # Resize modes keep the original file type
                save_format = ext.lstrip(".").upper()
                save_format = "JPEG" if save_format in ("JPG", "JPEG") \
                    else save_format
                if save_format not in ("PNG", "JPEG", "WEBP", "BMP",
                                       "TIFF", "GIF"):
                    save_format = "PNG"
                out_ext = _ext_for(save_format) if save_format == "JPEG" \
                    else ext.lower()
                out_path = os.path.join(folder, f"{base}_resized{out_ext}")

                if _kind == "max_side":
                    img.thumbnail((value, value), Image.LANCZOS)
                else:  # scale
                    img = img.resize(
                        (max(1, img.width * value // 100),
                         max(1, img.height * value // 100)), Image.LANCZOS)

                img = _save_mode(img, save_format)
                img.save(out_path, save_format)
                log(f"Resized: {filename} -> {os.path.basename(out_path)} "
                    f"({img.width}x{img.height})")
                done += 1
        except Exception as e:
            log(f"Failed: {filename}: {e}")
            failed += 1

    log(f"\nDone: {done} processed"
        + (f", {failed} failed" if failed else "")
        + (f", {skipped} skipped" if skipped else ""))


if __name__ == "__main__":
    main()
