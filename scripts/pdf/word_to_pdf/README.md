# Word to PDF Batch Converter

Converts all Word documents (.doc, .docx) in a folder to PDF format.

## Requirements

- **Windows only** - Uses COM automation
- **Microsoft Word** must be installed

## Features

- Batch converts all .doc and .docx files
- Skips already-converted files (if PDF exists)
- Preserves document properties and bookmarks
- Retries failed conversions automatically
- Ignores Word temp files (~$)

## Usage

1. Place your Word documents in this folder
2. Run the script:

```bash
python script.py
```

3. PDF files are created alongside the originals

## Notes

- Close any open Word documents before running
- Large documents may take longer to convert
- Script will retry up to 3 times on failures
