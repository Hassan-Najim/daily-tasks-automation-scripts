# Everything to PDF

## WHAT

Combines every convertible file in the working folder into a single PDF:
Word, Excel, PowerPoint, existing PDFs and images of any supported type -
merged in name order into `combined.pdf`.

## NEEDS

- Word/Excel/PowerPoint/PDF/image files in this folder
- **Windows with Microsoft Office** for the Office files - Word documents
  export with Word, spreadsheets with Excel, decks with PowerPoint.
  Files whose app is not installed are skipped with a warning
  (PDFs and images always work, no Office needed)
- HEIC images need `pillow-heif` (installed with the top-level requirements)
- Python packages: pypdf, Pillow, reportlab, pywin32

## HOW

1. Put your files in the working folder (in the TUI: `D` to change it)
2. Select the script and press Enter or R

## OUTPUT

- `combined.pdf` - every file converted and merged in name order
  - Office files are exported through the matching Office app
  - images become pages at their native size (DPI-aware), rotated per
    their EXIF orientation
  - existing PDFs are appended as-is
- Files that could not be converted are listed at the end; the original
  files are never modified
