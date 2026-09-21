# Image Format Converter

Batch converts all images in a folder to a chosen format. Originals are kept.

## Features

- Pick a target format: PNG, JPG, BMP, GIF, TIFF or WEBP
- Accepts PNG, JPG, JPEG, BMP, GIF, TIF, TIFF, WEBP input
- Preserves transparency for formats that support alpha
- Flattens transparent images onto a white background for JPG
- Files already in the target format are skipped

## Usage

1. Place images in this folder
2. Run the script:

```bash
python script.py
```

3. Pick a target format from the numbered list

## Programmatic Use

```python
from script import convert_images

convert_images(folder_path, target_format='PNG')
```

## Supported Input Formats

- PNG (.png)
- JPEG (.jpg, .jpeg)
- BMP (.bmp)
- GIF (.gif)
- TIFF (.tif, .tiff)
- WebP (.webp)
