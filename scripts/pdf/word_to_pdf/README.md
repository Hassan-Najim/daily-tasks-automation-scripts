# Word to PDF Batch Converter

## WHAT

Batch converts every Word document (.doc, .docx) in the folder to PDF.

## NEEDS

- **Windows only** - uses COM automation
- **Microsoft Word** must be installed
- .doc/.docx files in this folder
- Python packages: pywin32

## HOW

1. Place your Word documents in this folder
2. Close any open Word documents
3. Run the script:

```bash
python script.py
```

## OUTPUT

- A PDF next to each document
- Already-converted files are skipped; failed conversions retried up to 3 times
- Word temp files (~$) are ignored
