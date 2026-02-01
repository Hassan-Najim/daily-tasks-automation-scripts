import os
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def images_to_pdf(folder_path, output_pdf):
    images = [f for f in os.listdir(folder_path) if f.lower().endswith(('png', 'jpg', 'jpeg', 'bmp', 'gif'))]
    images.sort()  # Sort images by name

    if not images:
        print("No images found in the folder.")
        return

    # Initialize PDF
    c = canvas.Canvas(output_pdf, pagesize=letter)
    width, height = letter

    for img in images:
        img_path = os.path.join(folder_path, img)
        print(f"Processing image: {img_path}")
        image = Image.open(img_path)
        aspect = image.width / float(image.height)
        
        if aspect > 1:
            # Landscape orientation
            img_width = width
            img_height = width / aspect
        else:
            # Portrait orientation
            img_height = height
            img_width = height * aspect

        # Center the image
        x = (width - img_width) / 2
        y = (height - img_height) / 2

        c.drawImage(img_path, x, y, img_width, img_height)
        c.showPage()

    c.save()
    print(f"PDF saved as {output_pdf}")

if __name__ == "__main__":
    folder_path = os.path.dirname(os.path.realpath(__file__))
    print(f"Folder path: {folder_path}")
    output_pdf = os.path.join(folder_path, "output.pdf")
    print(f"Output PDF path: {output_pdf}")
    images_to_pdf(folder_path, output_pdf)
    print("Finished processing")
