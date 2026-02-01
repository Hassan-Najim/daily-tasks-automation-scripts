from PIL import Image
import os

# Get the current directory
current_dir = os.getcwd()

# Supported input image formats
valid_extensions = ('.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.webp')

# Loop through files in the directory
for filename in os.listdir(current_dir):
    if filename.lower().endswith(valid_extensions):
        file_path = os.path.join(current_dir, filename)
        new_filename = os.path.splitext(filename)[0] + '.png'
        new_filepath = os.path.join(current_dir, new_filename)

        try:
            with Image.open(file_path) as img:
                img.convert("RGB").save(new_filepath, "PNG")
            os.remove(file_path)
            print(f"Converted and removed original: {filename} → {new_filename}")
        except Exception as e:
            print(f"Failed to convert {filename}: {e}")
