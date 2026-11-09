"""
Shared helpers for the automation scripts.

The scripts in scripts/<category>/<name>/ import this module so behavior
(file type lists, sorting, HEIC support, Office PDF export) stays identical
across the collection. It is traced into the exe because main_bundled.py
imports it.

How scripts import it (works for the TUI, `python main.py`, and direct
`python scripts/.../script.py` runs):

    try:
        import script_common as sc
    except ImportError:
        import os as _os, sys as _sys
        _sys.path.insert(0, _os.path.dirname(_os.path.dirname(
            _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))))
        import script_common as sc
"""

import re

# --- File type lists -------------------------------------------------------

IMAGE_EXTS = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff',
              '.webp')
HEIC_EXTS = ('.heic', '.heif')
PDF_EXTS = ('.pdf',)
WORD_EXTS = ('.docx', '.doc')
EXCEL_EXTS = ('.xlsx', '.xls')
POWERPOINT_EXTS = ('.pptx', '.ppt')

# --- Helpers ---------------------------------------------------------------


def natural_key(name):
    """Sort key that orders numbers inside names sensibly (2 before 10)."""
    return [int(part) if part.isdigit() else part.lower()
            for part in re.split(r'(\d+)', name)]


def register_heif():
    """Enable HEIC/HEIF support if pillow-heif is installed. True if active."""
    try:
        from pillow_heif import register_heif_opener
        register_heif_opener()
        return True
    except ImportError:
        return False


def office_export_pdf(src_path, pdf_path, prog_id, log=print):
    """Convert an Office document to PDF via COM. Raises on failure.

    prog_id: "Word.Application", "Excel.Application" or "PowerPoint.Application"
    """
    import os as _os
    import time as _time

    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError(
            "pywin32 is required for Office conversion "
            "(pip install pywin32, Windows only)") from e

    src_path = _os.path.abspath(src_path)
    pdf_path = _os.path.abspath(pdf_path)
    app = None
    doc = None
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx(prog_id)
        # PowerPoint must stay visible (it rejects Visible=False); its
        # presentations are opened with WithWindow=False instead.
        if prog_id == "Word.Application":
            app.Visible = False
            app.ScreenUpdating = False
            app.DisplayAlerts = 0  # wdAlertsNone
        elif prog_id == "Excel.Application":
            app.Visible = False
            app.DisplayAlerts = False

        last_error = None
        for attempt in range(1, 4):
            try:
                if prog_id == "Excel.Application":
                    doc = app.Workbooks.Open(src_path)
                    doc.ExportAsFixedFormat(0, pdf_path)  # 0 = xlTypePDF
                elif prog_id == "PowerPoint.Application":
                    # Open(FileName, ReadOnly, Untitled, WithWindow)
                    doc = app.Presentations.Open(src_path, True, False, False)
                    doc.SaveAs(pdf_path, 32)  # 32 = ppSaveAsPDF
                else:
                    doc = app.Documents.Open(src_path)
                    doc.ExportAsFixedFormat(
                        OutputFileName=pdf_path,
                        ExportFormat=17,  # 17 = wdExportFormatPDF
                        OpenAfterExport=False,
                        OptimizeFor=0,
                        Range=0,
                        Item=0,
                        IncludeDocProps=True,
                        KeepIRM=True,
                        CreateBookmarks=0,
                        DocStructureTags=True,
                        BitmapMissingFonts=True,
                        UseISO19005_1=False,
                    )
                return pdf_path
            except Exception as e:
                last_error = e
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                    doc = None
                if attempt < 3:
                    _time.sleep(1.2 * attempt)
        raise RuntimeError(f"COM export failed after 3 attempts: {last_error}")
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
