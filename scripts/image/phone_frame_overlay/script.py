"""
Phone Frame Overlay
Composites screenshots onto a device frame mockup image.
"""

import os
from PIL import Image, ImageDraw


def create_placeholder_frame(frame_path: str) -> None:
    """Create a simple placeholder phone frame (dark rounded outline)."""
    width, height = 640, 1280
    frame = Image.new('RGBA', (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(frame)
    draw.rounded_rectangle([8, 8, width - 9, height - 9], radius=56,
                           outline=(24, 24, 24, 255), width=16)
    frame.save(frame_path)


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

    if not os.path.exists(screens_folder):
        os.makedirs(screens_folder)
        print("Created: screens/ folder")

    if not os.path.exists(frame_path):
        create_placeholder_frame(frame_path)
        print("Created: frame.png (placeholder phone outline)")
        print("Tip: replace frame.png with any device mockup PNG (with transparency).")

    print(f"Frame: {frame_path}")
    print(f"Screens: {screens_folder}")
    print(f"Output: {output_folder}")
    print()

    screens = [f for f in os.listdir(screens_folder)
               if f.lower().endswith('.png') and not f.startswith('.')]
    if not screens:
        print("No screenshots yet.")
        print(f"Drop PNG screenshots into: {screens_folder}")
        print("Then run Phone Frame Overlay again.")
        return

    overlay_frames(frame_path, screens_folder, output_folder)
    print("Done!")


if __name__ == "__main__":
    main()
