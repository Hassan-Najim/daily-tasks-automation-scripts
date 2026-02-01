# batch_word_to_pdf.py
import os
import time
from pathlib import Path

def list_word_files(folder: Path):
    return [
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in (".docx", ".doc")
        and not p.name.startswith("~$")     # skip Word lock/temp files
    ]

def convert_all_in_folder(folder: Path):
    try:
        import pythoncom
        import win32com.client
    except Exception as e:
        print("❌ Missing dependency. Install with:  pip install pywin32")
        raise

    files = list_word_files(folder)
    if not files:
        print("No .doc or .docx files found next to this script.")
        return

    # Initialize COM and Word once
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
                # Skip if already converted
                continue

            print(f"Converting: {doc_path.name} -> {pdf_path.name}")
            # Retry a couple of times for COM hiccups or locked files
            success = False
            for attempt in range(1, 4):
                try:
                    doc = word.Documents.Open(str(doc_path))
                    # ExportAsFixedFormat = best fidelity for PDF
                    # 17 = wdExportFormatPDF
                    doc.ExportAsFixedFormat(
                        OutputFileName=str(pdf_path),
                        ExportFormat=17,
                        OpenAfterExport=False,
                        OptimizeFor=0,          # 0=Print, 1=OnScreen
                        Range=0,                # 0=All document
                        From=1, To=1,           # ignored when Range=All
                        Item=0,                 # 0=DocumentContent
                        IncludeDocProps=True,
                        KeepIRM=True,
                        CreateBookmarks=1,      # 1=From headings
                        DocStructureTags=True,
                        BitmapMissingFonts=True,
                        UseISO19005_1=False
                    )
                    doc.Close(False)
                    success = True
                    converted += 1
                    break
                except Exception as e:
                    # try to close if partially opened
                    try:
                        doc.Close(False)
                    except:
                        pass
                    if attempt < 3:
                        time.sleep(1.2 * attempt)  # brief backoff
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

    print("\n— Summary —")
    print(f"✅ Converted: {converted}")
    if failures:
        print(f"❌ Failed ({len(failures)}):")
        for name, err in failures:
            print(f"   - {name}: {err}")
    else:
        print("🎉 No failures.")

if __name__ == "__main__":
    here = Path(__file__).resolve().parent
    convert_all_in_folder(here)
