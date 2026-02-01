# Images to PDF Converter (Custom Aspect Ratio)

Converts images to PDF preserving their native aspect ratio and DPI settings.

## Features

- Preserves original image aspect ratio (page size matches image)
- Respects embedded DPI metadata
- Auto-rotates based on EXIF orientation
- Multiple sort options: natural, modification time, or EXIF date
- Supports PNG, JPG, JPEG, BMP, GIF, TIF, TIFF formats

## Usage

1. Place your images in this folder
2. Run the script:

```bash
python script.py
```

3. Find `output.pdf` in the same folder

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

## Output

- Creates `output.pdf` with each page sized to match the image dimensions
- Ideal for documents, screenshots, or scanned pages where exact sizing matters
