# Image Format Converter

Batch converts images to PNG format.

## Features

- Converts JPG, JPEG, BMP, GIF, TIFF, WEBP to PNG
- Optionally deletes original files after conversion
- Preserves transparency where possible

## Usage

1. Place images in this folder
2. Run the script:

```bash
python script.py
```

3. Original files are replaced with PNG versions

## Configuration

To keep original files, edit `main()`:

```python
convert_to_png(folder_path, delete_originals=False)
```

## Supported Input Formats

- JPEG (.jpg, .jpeg)
- BMP (.bmp)
- GIF (.gif)
- TIFF (.tiff)
- WebP (.webp)
