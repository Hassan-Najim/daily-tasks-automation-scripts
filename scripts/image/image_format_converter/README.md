# Image Format Converter

## WHAT

Batch converts all images in the folder to a format you pick. Originals are kept.

## NEEDS

- Images (PNG, JPG, JPEG, BMP, GIF, TIF, TIFF, WEBP) in this folder
- Python packages: Pillow

## HOW

1. Place images in this folder
2. Run the script:

```bash
python script.py
```

3. Pick a target format from the numbered list (PNG, JPG, BMP, GIF, TIFF, WEBP)

## OUTPUT

- A converted copy of each image in this folder
- Originals are kept; files already in the target format are skipped
- Transparency is preserved; transparent images are flattened onto white for JPG

## Programmatic Use

```python
from script import convert_images

convert_images(folder_path, target_format='PNG')
```
