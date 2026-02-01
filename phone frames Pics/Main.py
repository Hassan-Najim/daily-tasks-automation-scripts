from PIL import Image
import os

# Paths
frame_path = 'frame.png'
screens_folder = 'screens'
output_folder = 'output'

os.makedirs(output_folder, exist_ok=True)

# Load frame
frame = Image.open(frame_path).convert('RGBA')
frame_width, frame_height = frame.size

for filename in os.listdir(screens_folder):
    if filename.lower().endswith('.png') and not filename.startswith('.'):
        screen_path = os.path.join(screens_folder, filename)
        try:
            screen = Image.open(screen_path).convert('RGBA')
        except Exception as e:
            print(f"Skipping {filename}: {e}")
            continue

        # Create blank transparent canvas the size of the frame
        canvas = Image.new('RGBA', (frame_width, frame_height), (0, 0, 0, 0))

        # Center the screen image on the canvas
        screen_x = (frame_width - screen.width) // 2
        screen_y = (frame_height - screen.height) // 2
        canvas.paste(screen, (screen_x, screen_y), screen)

        # Overlay the frame on top
        final_image = Image.alpha_composite(canvas, frame)

        # Save it
        output_path = os.path.join(output_folder, filename)
        final_image.save(output_path, format='PNG')
