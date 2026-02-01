"""
Phone Frame Overlay
Composites screenshots onto a device frame mockup image.
"""

import os
from PIL import Image


def overlay_frames(frame_path: str, screens_folder: str, output_folder: str) -> None:
    """
    Overlay screenshots onto a phone frame.
    
    Args:
        frame_path: Path to the frame PNG (with transparency)
        screens_folder: Folder containing screenshot images
        output_folder: Folder to save composited images
    """
    os.makedirs(output_folder, exist_ok=True)

    # Load frame
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

            # Create transparent canvas
            canvas = Image.new('RGBA', (frame_width, frame_height), (0, 0, 0, 0))

            # Center the screen on canvas
            screen_x = (frame_width - screen.width) // 2
            screen_y = (frame_height - screen.height) // 2
            canvas.paste(screen, (screen_x, screen_y), screen)

            # Overlay frame on top
            final_image = Image.alpha_composite(canvas, frame)

            # Save
            output_path = os.path.join(output_folder, filename)
            final_image.save(output_path, format='PNG')
            print(f"Created: {filename}")
            processed += 1

    print(f"\nProcessed {processed} images")


def main():
    """Run the overlay on default folders."""
    script_dir = os.path.dirname(os.path.realpath(__file__))
    
    frame_path = os.path.join(script_dir, 'frame.png')
    screens_folder = os.path.join(script_dir, 'screens')
    output_folder = os.path.join(script_dir, 'output')

    if not os.path.exists(frame_path):
        print(f"Error: frame.png not found in {script_dir}")
        print("Please add a frame.png file (device mockup with transparency)")
        return

    if not os.path.exists(screens_folder):
        print(f"Error: 'screens' folder not found in {script_dir}")
        print("Please create a 'screens' folder and add your screenshots")
        return

    print(f"Frame: {frame_path}")
    print(f"Screens: {screens_folder}")
    print(f"Output: {output_folder}")
    print()

    overlay_frames(frame_path, screens_folder, output_folder)
    print("Done!")


if __name__ == "__main__":
    main()
