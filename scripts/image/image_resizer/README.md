# Image Resizer

## WHAT

Batch-resizes or compresses the images in the working folder: cap the
longest side (1920/1280 px), scale to 50%, or create JPG copies under a
target file size. Originals are always kept.

## NEEDS

- Images (PNG, JPG, BMP, GIF, TIFF, WEBP, HEIC) in this folder
- HEIC images need `pillow-heif` (installed with the top-level requirements)
- Python packages: Pillow

## HOW

1. Put your images in the working folder
2. Run the script and pick a preset:

```bash
python script.py
```

Presets:

- Longest side -> 1920 px / 1280 px (keeps the original file type)
- Scale to 50 %
- JPG copies under 1 MB / 500 KB (iterative quality search; photos
  bigger than the target at minimum quality are shrunk further)

## OUTPUT

- A resized or compressed copy of each image: `<name>_resized.<ext>`
  (resize modes keep the original file type; target-size mode saves JPG)
- A per-image log line with the new dimensions or size; originals untouched
