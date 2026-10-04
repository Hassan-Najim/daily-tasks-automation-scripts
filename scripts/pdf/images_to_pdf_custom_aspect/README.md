# Images to PDF (Custom Aspect)

## WHAT

Combines images into a single PDF where each page keeps the image's real size and aspect ratio.

## NEEDS

- Images (PNG, JPG, JPEG, BMP, GIF, TIF, TIFF) in this folder
- Python packages: Pillow, reportlab

## HOW

1. Place your images in this folder
2. Run the script:

```bash
python script.py
```

## OUTPUT

- `output.pdf` with each page sized to match the image dimensions
- Respects embedded DPI metadata and EXIF rotation
- Ordered naturally (img2 before img10)

## Configuration

Edit the `main()` function to change sorting:

```python
# Sort by filename naturally (default) - Q2.png before Q10.png
images_to_pdf(folder_path, output_pdf, order_by="natural")

# Sort by file modification time
images_to_pdf(folder_path, output_pdf, order_by="mtime")

# Sort by EXIF date taken
images_to_pdf(folder_path, output_pdf, order_by="exif")

# Reverse any sort order
images_to_pdf(folder_path, output_pdf, order_by="natural", reverse=True)
```
