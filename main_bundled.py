"""
Automation Scripts - Bundled Version
All scripts embedded for standalone executable compilation.

This file is auto-usable but designed to be compiled with PyInstaller.
"""

import os
import sys
import tempfile
from pathlib import Path


# =============================================================================
# EMBEDDED SCRIPTS
# =============================================================================

# -----------------------------------------------------------------------------
# PDF: Images to PDF
# -----------------------------------------------------------------------------
def images_to_pdf_main():
    from PIL import Image
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter

    folder_path = os.getcwd()
    output_pdf = os.path.join(folder_path, "output.pdf")

    extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
    images = [f for f in os.listdir(folder_path) if f.lower().endswith(extensions)]
    images.sort()

    if not images:
        print("No images found in the current folder.")
        return

    c = canvas.Canvas(output_pdf, pagesize=letter)
    width, height = letter

    for img in images:
        img_path = os.path.join(folder_path, img)
        print(f"Processing: {img}")
        
        image = Image.open(img_path)
        aspect = image.width / float(image.height)

        if aspect > 1:
            img_width = width
            img_height = width / aspect
        else:
            img_height = height
            img_width = height * aspect

        x = (width - img_width) / 2
        y = (height - img_height) / 2

        c.drawImage(img_path, x, y, img_width, img_height)
        c.showPage()

    c.save()
    print(f"\nPDF saved: {output_pdf}")


# -----------------------------------------------------------------------------
# PDF: Images to PDF (Custom Aspect)
# -----------------------------------------------------------------------------
def images_to_pdf_custom_main():
    import re
    from datetime import datetime
    from PIL import Image, ImageOps, ExifTags
    from reportlab.pdfgen import canvas

    def natural_key(s):
        return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]

    folder_path = os.getcwd()
    output_pdf = os.path.join(folder_path, "output.pdf")
    default_dpi = 72

    extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff')
    images = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
              if f.lower().endswith(extensions)]

    if not images:
        print("No images found in the current folder.")
        return

    images.sort(key=lambda p: natural_key(os.path.basename(p)))
    c = canvas.Canvas(output_pdf)

    for img_path in images:
        print(f"Processing: {os.path.basename(img_path)}")
        
        with Image.open(img_path) as im:
            im = ImageOps.exif_transpose(im)
            px_w, px_h = im.size
            dpi_x, dpi_y = im.info.get("dpi", (default_dpi, default_dpi))

            page_w = px_w * 72.0 / (dpi_x or default_dpi)
            page_h = px_h * 72.0 / (dpi_y or default_dpi)

            c.setPageSize((page_w, page_h))
            c.drawImage(img_path, 0, 0, width=page_w, height=page_h,
                        preserveAspectRatio=False, mask='auto')
            c.showPage()

    c.save()
    print(f"\nPDF saved: {output_pdf}")


# -----------------------------------------------------------------------------
# PDF: Word to PDF
# -----------------------------------------------------------------------------
def word_to_pdf_main():
    import time
    
    if sys.platform != "win32":
        print("Error: Word to PDF only works on Windows with Microsoft Word installed.")
        return

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("Error: pywin32 not installed. Run: pip install pywin32")
        return

    folder = Path(os.getcwd())
    
    files = [
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in (".docx", ".doc")
        and not p.name.startswith("~$")
    ]

    if not files:
        print("No .doc or .docx files found in current folder.")
        return

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.ScreenUpdating = False
    word.DisplayAlerts = 0

    failures = []
    converted = 0

    try:
        for doc_path in files:
            pdf_path = doc_path.with_suffix(".pdf")
            
            if pdf_path.exists():
                print(f"Skipping (exists): {doc_path.name}")
                continue

            print(f"Converting: {doc_path.name}")
            
            success = False
            for attempt in range(1, 4):
                try:
                    doc = word.Documents.Open(str(doc_path))
                    doc.ExportAsFixedFormat(
                        OutputFileName=str(pdf_path),
                        ExportFormat=17,
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
                print(f"  FAILED: {doc_path.name}")
    finally:
        word.Quit()
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    print(f"\nConverted: {converted}")
    if failures:
        print(f"Failed: {len(failures)}")
        for name, err in failures:
            print(f"  - {name}: {err}")


# -----------------------------------------------------------------------------
# Image: Phone Frame Overlay
# -----------------------------------------------------------------------------
def phone_frame_overlay_main():
    from PIL import Image

    folder = os.getcwd()
    frame_path = os.path.join(folder, 'frame.png')
    screens_folder = os.path.join(folder, 'screens')
    output_folder = os.path.join(folder, 'output')

    if not os.path.exists(frame_path):
        print(f"Error: frame.png not found in {folder}")
        print("Please add a frame.png file (device mockup with transparency)")
        return

    if not os.path.exists(screens_folder):
        print(f"Error: 'screens' folder not found in {folder}")
        print("Please create a 'screens' folder and add your screenshots")
        return

    os.makedirs(output_folder, exist_ok=True)

    frame = Image.open(frame_path).convert('RGBA')
    frame_width, frame_height = frame.size

    processed = 0
    for filename in os.listdir(screens_folder):
        if filename.lower().endswith('.png') and not filename.startswith('.'):
            screen_path = os.path.join(screens_folder, filename)
            
            try:
                screen = Image.open(screen_path).convert('RGBA')
            except Exception as e:
                print(f"Skipping {filename}: {e}")
                continue

            canvas = Image.new('RGBA', (frame_width, frame_height), (0, 0, 0, 0))
            screen_x = (frame_width - screen.width) // 2
            screen_y = (frame_height - screen.height) // 2
            canvas.paste(screen, (screen_x, screen_y), screen)

            final_image = Image.alpha_composite(canvas, frame)

            output_path = os.path.join(output_folder, filename)
            final_image.save(output_path, format='PNG')
            print(f"Created: {filename}")
            processed += 1

    print(f"\nProcessed {processed} images -> output/")


# -----------------------------------------------------------------------------
# Image: Format Converter
# -----------------------------------------------------------------------------
def image_format_converter_main():
    from PIL import Image

    folder_path = os.getcwd()
    supported = ('.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.webp')

    converted = 0
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(supported):
            file_path = os.path.join(folder_path, filename)
            new_filename = os.path.splitext(filename)[0] + '.png'
            new_filepath = os.path.join(folder_path, new_filename)

            try:
                with Image.open(file_path) as img:
                    if img.mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGBA')
                    else:
                        img = img.convert('RGB')
                    img.save(new_filepath, "PNG")
                
                os.remove(file_path)
                print(f"Converted: {filename} -> {new_filename}")
                converted += 1
            except Exception as e:
                print(f"Failed: {filename}: {e}")

    print(f"\nConverted {converted} images to PNG")


# -----------------------------------------------------------------------------
# Document: Word Find & Replace
# -----------------------------------------------------------------------------
def word_find_replace_main():
    import json
    from docx import Document

    folder = os.getcwd()
    
    # Look for corrections.json
    json_path = os.path.join(folder, "corrections.json")
    if not os.path.exists(json_path):
        print("Error: corrections.json not found in current folder")
        print("\nCreate a corrections.json file with format:")
        print('[')
        print('  {"OgSentence": "original text", "NewSentence": "replacement"},')
        print('  ...')
        print(']')
        return

    # Find .docx file
    docx_files = [f for f in os.listdir(folder) if f.endswith('.docx') and not f.startswith('~$')]
    if not docx_files:
        print("Error: No .docx file found in current folder")
        return
    
    if len(docx_files) > 1:
        print("Multiple .docx files found. Select one:")
        for i, f in enumerate(docx_files, 1):
            print(f"  {i}. {f}")
        choice = input("Enter number: ").strip()
        try:
            doc_file = docx_files[int(choice) - 1]
        except:
            print("Invalid choice")
            return
    else:
        doc_file = docx_files[0]

    doc_path = os.path.join(folder, doc_file)

    # Load corrections
    with open(json_path, 'r', encoding='utf-8') as f:
        corrections = json.load(f)

    print(f"Document: {doc_file}")
    print(f"Corrections: {len(corrections)}")

    # Process document
    document = Document(doc_path)
    replacements = 0

    for entry in corrections:
        original = entry.get("OgSentence")
        replacement = entry.get("NewSentence")

        if not original or not replacement:
            continue

        for paragraph in document.paragraphs:
            if original in paragraph.text:
                paragraph.text = paragraph.text.replace(original, replacement, 1)
                replacements += 1
                break

        for table in document.tables:
            found = False
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if original in paragraph.text:
                            paragraph.text = paragraph.text.replace(original, replacement, 1)
                            replacements += 1
                            found = True
                            break
                    if found:
                        break
                if found:
                    break
            if found:
                break

    output_path = doc_path.replace(".docx", "_updated.docx")
    document.save(output_path)
    print(f"\nReplacements: {replacements}")
    print(f"Saved: {output_path}")


# =============================================================================
# MAIN MENU
# =============================================================================

SCRIPTS = {
    "pdf": [
        ("Images to PDF", images_to_pdf_main),
        ("Images to PDF (Custom Aspect)", images_to_pdf_custom_main),
        ("Word to PDF", word_to_pdf_main),
    ],
    "image": [
        ("Phone Frame Overlay", phone_frame_overlay_main),
        ("Image Format Converter", image_format_converter_main),
    ],
    "document": [
        ("Word Find & Replace", word_find_replace_main),
    ],
}

CATEGORY_NAMES = {
    "pdf": "PDF",
    "image": "Image",
    "document": "Document",
}


def display_menu():
    print("\n" + "="*50)
    print("   AUTOMATION SCRIPTS")
    print("="*50)
    print(f"\n   Working directory: {os.getcwd()}")

    index = 1
    script_map = {}

    for category, scripts in SCRIPTS.items():
        print(f"\n[{CATEGORY_NAMES[category]}]")
        for name, func in scripts:
            print(f"  {index}. {name}")
            script_map[index] = (name, func)
            index += 1

    print(f"\n  c. Change directory")
    print(f"  q. Quit")
    print("="*50)

    return script_map


def change_directory():
    new_dir = input("Enter path: ").strip()
    if os.path.isdir(new_dir):
        os.chdir(new_dir)
        print(f"Changed to: {os.getcwd()}")
    else:
        print("Invalid directory")


def main():
    print("\n" + "="*50)
    print("   AUTOMATION SCRIPTS - Standalone Edition")
    print("="*50)
    
    while True:
        script_map = display_menu()

        choice = input("\nEnter choice: ").strip().lower()

        if choice == 'q':
            print("\nGoodbye!")
            break
        elif choice == 'c':
            change_directory()
            continue

        try:
            choice_num = int(choice)
            if choice_num in script_map:
                name, func = script_map[choice_num]
                print(f"\n{'='*50}")
                print(f"Running: {name}")
                print(f"Directory: {os.getcwd()}")
                print('='*50 + "\n")
                try:
                    func()
                except Exception as e:
                    print(f"\nError: {e}")
                input("\nPress Enter to continue...")
            else:
                print("Invalid choice.")
        except ValueError:
            print("Invalid input.")


# =============================================================================
# BUILD FUNCTION
# =============================================================================

def build_exe():
    """Compile this script into a standalone executable."""
    import subprocess
    
    # Check/install PyInstaller
    try:
        import PyInstaller
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    script_path = Path(__file__).resolve()
    base_dir = script_path.parent
    
    print("="*50)
    print("Building Automation Scripts Executable")
    print("="*50)
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--name", "automation_scripts",
        "--console",
        "--clean",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--hidden-import", "PIL.ImageOps",
        "--hidden-import", "PIL.ExifTags",
        "--hidden-import", "reportlab",
        "--hidden-import", "reportlab.pdfgen",
        "--hidden-import", "reportlab.pdfgen.canvas",
        "--hidden-import", "reportlab.lib.pagesizes",
        "--hidden-import", "docx",
        "--hidden-import", "docx.document",
        "--collect-submodules", "reportlab",
        "--collect-submodules", "docx",
        str(script_path)
    ]
    
    if sys.platform == "win32":
        cmd.extend([
            "--hidden-import", "win32com",
            "--hidden-import", "win32com.client",
            "--hidden-import", "pythoncom",
        ])
    
    print("\nRunning PyInstaller...\n")
    result = subprocess.run(cmd, cwd=base_dir)
    
    if result.returncode == 0:
        exe_name = "automation_scripts.exe" if sys.platform == "win32" else "automation_scripts"
        exe_path = base_dir / "dist" / exe_name
        print("\n" + "="*50)
        print("BUILD SUCCESSFUL!")
        print("="*50)
        print(f"\nExecutable: {exe_path}")
        if exe_path.exists():
            print(f"Size: {exe_path.stat().st_size / 1024 / 1024:.1f} MB")
        print("\nCopy this file anywhere and run it.")
    else:
        print("\nBuild failed!")
        sys.exit(1)


if __name__ == "__main__":
    if "--build" in sys.argv:
        build_exe()
    else:
        main()
