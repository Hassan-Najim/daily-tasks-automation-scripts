"""
Image Format Converter
Batch converts images to PNG format and removes the originals.
"""

import os
from PIL import Image


SUPPORTED_FORMATS = ('.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.webp')


def convert_to_png(folder_path: str, delete_originals: bool = True) -> None:
    """
    Convert all supported images in folder to PNG.
    
    Args:
        folder_path: Directory containing images
        delete_originals: Whether to delete original files after conversion
    """
    converted = 0
    failed = 0

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(SUPPORTED_FORMATS):
            file_path = os.path.join(folder_path, filename)
            new_filename = os.path.splitext(filename)[0] + '.png'
            new_filepath = os.path.join(folder_path, new_filename)

            try:
                with Image.open(file_path) as img:
                    # Convert to RGB if necessary (some formats need this)
                    if img.mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGBA')
                    else:
                        img = img.convert('RGB')
                    
                    img.save(new_filepath, "PNG")
                
                if delete_originals:
                    os.remove(file_path)
                    print(f"Converted and removed: {filename} -> {new_filename}")
                else:
                    print(f"Converted: {filename} -> {new_filename}")
                
                converted += 1
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")
                failed += 1

    print(f"\nConverted: {converted}")
    if failed:
        print(f"Failed: {failed}")


def main():
    """Run the converter on the current directory."""
    folder_path = os.path.dirname(os.path.realpath(__file__))
    
    print(f"Folder: {folder_path}")
    print(f"Supported formats: {', '.join(SUPPORTED_FORMATS)}")
    print()
    
    convert_to_png(folder_path, delete_originals=True)
    print("Done!")


if __name__ == "__main__":
    main()
