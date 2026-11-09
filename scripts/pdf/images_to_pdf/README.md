# Images to PDF

## WHAT

Combines every image in the folder into a single PDF file.

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

- `output.pdf` with one image per page
- Images are centered and scaled to fit letter size (8.5" x 11")
- Photos are rotated per their EXIF orientation (phone pictures come out upright)
- Ordered naturally (img2 before img10); animated GIFs use their first frame
