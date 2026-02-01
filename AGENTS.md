# AGENTS.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview

A collection of standalone Python automation scripts for daily tasks including:
- **Image to PDF conversion** - Converting images to PDFs with aspect ratio preservation
- **Phone frame overlays** - Compositing screenshots onto device frame mockups
- **Image format conversion** - Batch converting images to PNG
- **Word to PDF conversion** - Batch converting .docx/.doc files to PDF (Windows only, requires MS Word)
- **Grammar fix/find-replace** - Automated find-and-replace in Word documents using JSON-formatted corrections

## Project Structure

Each script is self-contained in its own folder:
- `images to pdf project/` - Image-to-PDF converters
- `phone frames Pics/` - Phone frame overlay tool with `Screens/` subfolder for image conversion
- `words to pdf/` - Word document to PDF batch converter
- `Grammer Fix Scrpt/` - Word document find-and-replace utility

## Dependencies

| Script | Required Packages |
|--------|-------------------|
| Image to PDF | `Pillow`, `reportlab` |
| Phone frames | `Pillow` |
| Word to PDF | `pywin32` (Windows only) |
| Grammar fix | `python-docx` |

Install with: `pip install Pillow reportlab python-docx pywin32`

## Running Scripts

Scripts are designed to run from their respective directories:
```powershell
# Image to PDF (place images in same folder as script)
python "images to pdf project/images_to_pdf.py"

# Phone frame overlay (requires frame.png and screens/ subfolder)
cd "phone frames Pics"
python Main.py

# Word to PDF (place .docx files in same folder as script)
python "words to pdf/sample test/main.py"
```

## Key Patterns

- Scripts use `os.path.dirname(os.path.realpath(__file__))` to process files relative to the script location
- Image processing scripts automatically find supported image formats (png, jpg, jpeg, bmp, gif, tiff, webp)
- Output files are created alongside inputs or in designated `output/` subfolders
- Word to PDF converter uses COM automation and requires Microsoft Word installed
