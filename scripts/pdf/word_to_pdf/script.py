"""
Word to PDF Batch Converter
Converts all .doc and .docx files in a folder to PDF.
Windows only - requires Microsoft Word installed.
"""

import os
import time
from pathlib import Path


def list_word_files(folder: Path) -> list:
    """List all Word files in folder, excluding temp files."""
    return [
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in (".docx", ".doc")
        and not p.name.startswith("~$")
    ]


def convert_all_in_folder(folder: Path) -> None:
    """
    Convert all Word documents in folder to PDF.
    
    Args:
        folder: Directory containing Word files
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("❌ Missing dependency. Install with: pip install pywin32")
        raise

    files = list_word_files(folder)
    if not files:
        print("No .doc or .docx files found.")
        return

    # Initialize COM and Word
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.ScreenUpdating = False
    word.DisplayAlerts = 0  # wdAlertsNone

    failures = []
    converted = 0
    
    try:
        for doc_path in files:
            pdf_path = doc_path.with_suffix(".pdf")
            
            if pdf_path.exists():
                print(f"Skipping (already exists): {doc_path.name}")
                continue

            print(f"Converting: {doc_path.name}")
            
            success = False
            for attempt in range(1, 4):
                try:
                    doc = word.Documents.Open(str(doc_path))
                    doc.ExportAsFixedFormat(
                        OutputFileName=str(pdf_path),
                        ExportFormat=17,  # wdExportFormatPDF
                        OpenAfterExport=False,
                        OptimizeFor=0,
                        Range=0,
                        From=1, To=1,
                        Item=0,
                        IncludeDocProps=True,
                        KeepIRM=True,
                        CreateBookmarks=1,
                        DocStructureTags=True,
                        BitmapMissingFonts=True,
                        UseISO19005_1=False
                    )
                    doc.Close(False)
                    success = True
                    converted += 1
                    break
                except Exception as e:
                    try:
                        doc.Close(False)
                    except:
                        pass
                    if attempt < 3:
                        time.sleep(1.2 * attempt)
                    else:
                        failures.append((doc_path.name, str(e)))
            
            if not success:
                print(f"  ❌ FAILED: {doc_path.name}")
    finally:
        word.Quit()
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    # Summary
    print("\n— Summary —")
    print(f"✅ Converted: {converted}")
    if failures:
        print(f"❌ Failed ({len(failures)}):")
        for name, err in failures:
            print(f"   - {name}: {err}")
    else:
        print("🎉 No failures.")


def main():
    """Run the converter on the current directory."""
    folder = Path(__file__).resolve().parent
    print(f"Folder: {folder}")
    convert_all_in_folder(folder)


if __name__ == "__main__":
    main()
