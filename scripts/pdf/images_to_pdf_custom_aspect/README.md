# Images to PDF (Custom Aspect)

## WHAT

Combines images into a single PDF where each page keeps the image's real size and aspect ratio.

## NEEDS

- Images (PNG, JPG, BMP, GIF, TIFF, WEBP, HEIC) in this folder
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

## Page order

When run from the TUI you pick the order from a list (press Esc for the
default); standalone runs default to natural order:

- Natural (1, 2, 10) - Q2.png before Q10.png
- Name (A-Z)
- Date modified
- EXIF date taken
