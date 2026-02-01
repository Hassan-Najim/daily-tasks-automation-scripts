# Automation Scripts Collection

A collection of Python scripts to automate everyday tasks.

## Quick Start

```bash
# Install all dependencies
pip install -r requirements.txt

# Run the main menu
python main.py
```

## Scripts

| Category | Script | Description |
|----------|--------|-------------|
| [PDF](scripts/pdf/images_to_pdf) | Images to PDF | Convert images to a single PDF |
| [PDF](scripts/pdf/images_to_pdf_custom_aspect) | Images to PDF (Custom) | Convert with native aspect ratios |
| [PDF](scripts/pdf/word_to_pdf) | Word to PDF | Batch convert Word docs to PDF |
| [Image](scripts/image/phone_frame_overlay) | Phone Frame Overlay | Add device frames to screenshots |
| [Image](scripts/image/image_format_converter) | Image Converter | Batch convert images to PNG |
| [Document](scripts/document/word_find_replace) | Word Find & Replace | Bulk text replacement in Word docs |

## Project Structure

```
automation-scripts/
├── main.py                 # CLI menu to run any script
├── requirements.txt        # All dependencies
├── scripts/
│   ├── pdf/
│   │   ├── images_to_pdf/
│   │   ├── images_to_pdf_custom_aspect/
│   │   └── word_to_pdf/
│   ├── image/
│   │   ├── phone_frame_overlay/
│   │   └── image_format_converter/
│   └── document/
│       └── word_find_replace/
└── .gitignore
```

Each script folder contains:
- `script.py` - The main script
- `README.md` - Usage instructions
- `requirements.txt` - Script-specific dependencies

## Running Individual Scripts

You can run scripts directly:

```bash
cd scripts/pdf/images_to_pdf
pip install -r requirements.txt
python script.py
```

## Requirements

- Python 3.8+
- Windows (for Word to PDF script - requires MS Word)
