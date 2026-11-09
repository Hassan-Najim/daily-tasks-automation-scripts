# PDF Merge & Split

## WHAT

Merges every PDF in the working folder into one file, or splits one PDF
into single-page PDFs.

## NEEDS

- PDFs in this folder (2+ to merge, 1+ to split)
- Python packages: pypdf

## HOW

1. Put your PDFs in the working folder
2. Run the script and pick **Merge** or **Split**
3. For Split, pick which PDF when there are several

```bash
python script.py
```

## OUTPUT

- Merge: `merged.pdf` with every PDF combined in name order
  (`combined.pdf`/`merged.pdf` outputs from other scripts are excluded)
- Split: a `split/` folder with one PDF per page, named
  `<name>_page_001.pdf`, `<name>_page_002.pdf`, ...
- Encrypted or unreadable PDFs are reported and skipped; the originals
  are never modified
