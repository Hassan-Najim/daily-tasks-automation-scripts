"""
Image Format Converter
Batch converts all images in the working folder to a chosen format
(originals are kept). HEIC is supported when pillow-heif is installed.
"""

import os
import sys

try:
    import script_common as sc
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
    import script_common as sc

from PIL import Image

TARGET_FORMATS = ['PNG', 'JPG', 'BMP', 'GIF', 'TIFF', 'WEBP']
EXTS_BY_FORMAT = {'PNG': ('.png',), 'JPG': ('.jpg', '.jpeg'),
                  'BMP': ('.bmp',), 'GIF': ('.gif',),
                  'TIFF': ('.tif', '.tiff'), 'WEBP': ('.webp',),
                  'HEIC': ('.heic', '.heif')}


def convert_images(folder_path, target_format, log=print):
    """Convert all supported images in folder to the target format."""
    save_format = 'JPEG' if target_format == 'JPG' else target_format
    target_ext = EXTS_BY_FORMAT[target_format][0]

    converted = 0
    skipped = 0
    failed = 0

    filenames = sorted(
        (f for f in os.listdir(folder_path)
         if f.lower().endswith(sc.IMAGE_EXTS + sc.HEIC_EXTS)
         and not f.startswith("~$")),
        key=sc.natural_key)

    for filename in filenames:
        if filename.lower().endswith(EXTS_BY_FORMAT.get(target_format, ())):
            skipped += 1
            continue

        file_path = os.path.join(folder_path, filename)
        new_filename = os.path.splitext(filename)[0] + target_ext
        new_filepath = os.path.join(folder_path, new_filename)

        try:
            with Image.open(file_path) as img:
                img.load()
                if target_format == 'JPG':
                    if img.mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGBA')
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        background.paste(img, mask=img.split()[-1])
                        img = background
                    elif img.mode != 'RGB':
                        img = img.convert('RGB')
                elif target_format == 'GIF':
                    if img.mode not in ('P', 'L', 'RGB'):
                        img = img.convert('P')
                else:
                    if img.mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGBA')
                    elif img.mode != 'RGB':
                        img = img.convert('RGB')
                img.save(new_filepath, save_format)

            log(f"Converted: {filename} -> {new_filename}")
            converted += 1
        except Exception as e:
            log(f"Failed: {filename}: {e}")
            failed += 1

    log(f"\nConverted: {converted} image(s) to {target_format}")
    if skipped:
        log(f"Skipped: {skipped} already in target format")
    if failed:
        log(f"Failed: {failed}")
    return converted


def main(log=print, ask=None):
    folder_path = os.getcwd()

    heic = sc.register_heif()
    formats = TARGET_FORMATS + (['HEIC'] if heic else [])

    images = [f for f in os.listdir(folder_path)
              if f.lower().endswith(sc.IMAGE_EXTS + sc.HEIC_EXTS)]
    if not images:
        log("No supported images found in this folder.")
        return

    log(f"Folder: {folder_path}")
    log("Supported: "
        + ", ".join(ext.lstrip('.') for ext in sc.IMAGE_EXTS + sc.HEIC_EXTS))

    choice = None
    if ask is not None:
        choice = ask("Select target format", formats)
    else:
        log("Select target format:")
        for i, fmt in enumerate(formats, 1):
            log(f"  {i}. {fmt}")
        raw = input(f"Choice [1-{len(formats)}]: ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(formats):
            choice = formats[int(raw) - 1]
    if not choice:
        log("Cancelled.")
        return

    convert_images(folder_path, choice, log=log)


if __name__ == "__main__":
    main()
