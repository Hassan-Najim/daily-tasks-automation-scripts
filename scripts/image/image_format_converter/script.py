"""
Image Format Converter
Batch converts all images in a folder to a chosen format (originals kept).
"""

import os
from PIL import Image


SUPPORTED_FORMATS = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff', '.webp')
TARGET_FORMATS = ['PNG', 'JPG', 'BMP', 'GIF', 'TIFF', 'WEBP']
EXTS_BY_FORMAT = {'PNG': ('.png',), 'JPG': ('.jpg', '.jpeg'), 'BMP': ('.bmp',),
                  'GIF': ('.gif',), 'TIFF': ('.tif', '.tiff'), 'WEBP': ('.webp',)}


def convert_images(folder_path: str, target_format: str) -> None:
    """
    Convert all supported images in folder to the target format.

    Args:
        folder_path: Directory containing images
        target_format: One of TARGET_FORMATS (e.g. 'PNG', 'JPG')
    """
    save_format = 'JPEG' if target_format == 'JPG' else target_format
    target_ext = EXTS_BY_FORMAT[target_format][0]

    converted = 0
    skipped = 0
    failed = 0

    for filename in sorted(os.listdir(folder_path)):
        if not filename.lower().endswith(SUPPORTED_FORMATS):
            continue
        if filename.lower().endswith(EXTS_BY_FORMAT[target_format]):
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

            print(f"Converted: {filename} -> {new_filename}")
            converted += 1
        except Exception as e:
            print(f"Failed to convert {filename}: {e}")
            failed += 1

    print(f"\nConverted: {converted} image(s) to {target_format}")
    if skipped:
        print(f"Skipped: {skipped} already in target format")
    if failed:
        print(f"Failed: {failed}")


def main():
    """Run the converter on the current directory."""
    folder_path = os.path.dirname(os.path.realpath(__file__))

    print(f"Folder: {folder_path}")
    print(f"Supported formats: {', '.join(SUPPORTED_FORMATS)}")
    print()

    images = [f for f in os.listdir(folder_path)
              if f.lower().endswith(SUPPORTED_FORMATS)]
    if not images:
        print("No supported images found.")
        return

    print("Select target format:")
    for i, fmt in enumerate(TARGET_FORMATS, 1):
        print(f"  {i}. {fmt}")
    raw = input(f"Choice [1-{len(TARGET_FORMATS)}]: ").strip()
    if not (raw.isdigit() and 1 <= int(raw) <= len(TARGET_FORMATS)):
        print("Cancelled.")
        return
    target_format = TARGET_FORMATS[int(raw) - 1]

    convert_images(folder_path, target_format)
    print("Done!")


if __name__ == "__main__":
    main()
