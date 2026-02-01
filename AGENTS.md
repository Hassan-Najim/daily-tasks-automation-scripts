# AGENTS.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview

A collection of Python automation scripts organized by category. Each script is self-contained with its own README and requirements.

## Project Structure

```
scripts/
├── pdf/                          # PDF generation scripts
│   ├── images_to_pdf/            # Basic image-to-PDF converter
│   ├── images_to_pdf_custom_aspect/  # Preserves native aspect ratios
│   └── word_to_pdf/              # Word doc batch converter (Windows)
├── image/                        # Image processing scripts
│   ├── phone_frame_overlay/      # Device mockup compositor
│   └── image_format_converter/   # Batch PNG converter
└── document/                     # Document manipulation
    └── word_find_replace/        # Bulk find-replace in Word docs
```

## Running Scripts

### Via Main Menu
```bash
python main.py
```

### Individual Scripts
```bash
cd scripts/pdf/images_to_pdf
python script.py
```

## Dependencies

Install all: `pip install -r requirements.txt`

| Package | Used By |
|---------|--------|
| Pillow | All image scripts |
| reportlab | PDF scripts |
| python-docx | word_find_replace |
| pywin32 | word_to_pdf (Windows only) |

## Key Patterns

- All scripts have a `main()` function as entry point
- Scripts process files relative to their own directory via `os.path.dirname(os.path.realpath(__file__))`
- Output goes to `output/` subfolder or alongside input files
- Each script folder has: `script.py`, `README.md`, `requirements.txt`

## Adding New Scripts

1. Create folder under appropriate category: `scripts/<category>/<script_name>/`
2. Add `script.py` with a `main()` function
3. Add `README.md` and `requirements.txt`
4. Register in `main.py` SCRIPTS dict
