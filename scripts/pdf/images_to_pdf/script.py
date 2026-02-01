"""
Images to PDF Converter
Converts all images in a folder to a single PDF file with letter page size.
"""

import os
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter


def images_to_pdf(folder_path: str, output_pdf: str) -> None:
    """
    Convert all images in folder_path to a single PDF.
    
    Args:
        folder_path: Directory containing images
        output_pdf: Output PDF file path
    """
    extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
    images = [f for f in os.listdir(folder_path) if f.lower().endswith(extensions)]
    images.sort()

    if not images:
        print("No images found in the folder.")
        return

    c = canvas.Canvas(output_pdf, pagesize=letter)
    width, height = letter

    for img in images:
        img_path = os.path.join(folder_path, img)
        print(f"Processing: {img}")
        
        image = Image.open(img_path)
        aspect = image.width / float(image.height)

        if aspect > 1:
            # Landscape
            img_width = width
            img_height = width / aspect
        else:
            # Portrait
            img_height = height
            img_width = height * aspect

        # Center the image
        x = (width - img_width) / 2
        y = (height - img_height) / 2

        c.drawImage(img_path, x, y, img_width, img_height)
        c.showPage()

    c.save()
    print(f"PDF saved: {output_pdf}")


def main():
    """Run the converter on the current directory."""
    folder_path = os.path.dirname(os.path.realpath(__file__))
    output_pdf = os.path.join(folder_path, "output.pdf")
    
    print(f"Folder: {folder_path}")
    print(f"Output: {output_pdf}")
    
    images_to_pdf(folder_path, output_pdf)
    print("Done!")


if __name__ == "__main__":
    main()
