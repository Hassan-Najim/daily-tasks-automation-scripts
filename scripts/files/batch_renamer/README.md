# Batch Renamer

## WHAT

Renames all files in the working folder with one rule: sequence numbers,
lowercase-with-dashes, or a YYYYMMDD date-stamp prefix. Every change is
previewed first - nothing is renamed until you confirm.

## NEEDS

- Files in this folder
- No third-party Python packages (standard library only)

## HOW

1. Run the script and pick a rule:

```bash
python script.py
```

Rules:

- **Number the files**: `report.pdf` -> `report 01.pdf`,
  `report 02.pdf`, ... (files sharing a base name are numbered together)
- **Lowercase with dashes**: `My File (2).JPG` -> `my-file-2.jpg`
- **Date stamp**: `name.ext` -> `20251109_name.ext` (from the date modified)

2. Check the preview (old -> new) and confirm

## OUTPUT

- Renamed files in place; the preview shows every change and you can
  cancel before anything happens
- Name collisions get a numbered suffix; folders and hidden files are
  ignored
