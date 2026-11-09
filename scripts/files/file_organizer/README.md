# File Organizer

## WHAT

Sorts the loose files in the working folder into subfolders, either by
file type (Images, Documents, Archives, ...) or by year-month from the
date modified. Folders themselves are never touched.

## NEEDS

- Files in this folder
- No third-party Python packages (standard library only)

## HOW

1. Run the script and pick **By file type** or **By year-month**:

```bash
python script.py
```

Type folders: Images, Documents, Spreadsheets, Presentations, Archives,
Audio, Video, Code, Installers - anything else goes to `Other`.

## OUTPUT

- Subfolders like `Images/`, `Documents/` or `2025-01/` with the files
  moved inside
- Name collisions get a numbered suffix (`report.pdf`,
  `report_1.pdf`, ...); hidden files (`.…`) and Office temp files
  (`~$…`) are ignored
