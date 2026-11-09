"""
Phone Frame Overlay
Composites screenshots from a screens/ folder onto a device frame mockup
(frame.png) and saves the framed copies into output/. Oversized screenshots
are scaled down to fit inside the frame.
"""

import os
import sys

try:
    import script_common as sc
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
    import script_common as sc

from PIL import Image, ImageDraw

FRAME_NAME = 'frame.png'
SCREENS_DIR = 'screens'
OUTPUT_DIR = 'output'


def create_placeholder_frame(frame_path):
    """Create a simple placeholder phone frame (dark rounded outline)."""
    width, height = 640, 1280
    frame = Image.new('RGBA', (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(frame)
    draw.rounded_rectangle([8, 8, width - 9, height - 9], radius=56,
                           outline=(24, 24, 24, 255), width=16)
    frame.save(frame_path)


def overlay_frames(frame_path, screens_folder, output_folder, log=print):
    """Overlay screenshots from screens_folder onto the phone frame."""
    os.makedirs(output_folder, exist_ok=True)

    frame = Image.open(frame_path).convert('RGBA')
    frame_width, frame_height = frame.size

    screens = sorted(
        (f for f in os.listdir(screens_folder)
         if f.lower().endswith(sc.IMAGE_EXTS + sc.HEIC_EXTS)
         and not f.startswith('.')),
        key=sc.natural_key)

    processed = 0
    for filename in screens:
        screen_path = os.path.join(screens_folder, filename)
        try:
            screen = Image.open(screen_path).convert('RGBA')
        except Exception as e:
            log(f"Skipping {filename}: {e}")
            continue

        # Scale oversized screenshots down to fit inside the frame
        if screen.width > frame_width or screen.height > frame_height:
            screen.thumbnail((frame_width, frame_height), Image.LANCZOS)

        canvas = Image.new('RGBA', (frame_width, frame_height), (0, 0, 0, 0))
        screen_x = (frame_width - screen.width) // 2
        screen_y = (frame_height - screen.height) // 2
        canvas.paste(screen, (screen_x, screen_y), screen)

        final_image = Image.alpha_composite(canvas, frame)

        output_path = os.path.join(output_folder,
                                   os.path.splitext(filename)[0] + '.png')
        final_image.save(output_path, format='PNG')
        log(f"Created: {os.path.basename(output_path)}")
        processed += 1

    log(f"\nProcessed {processed} image(s) -> {OUTPUT_DIR}/")
    return processed


def main(log=print, ask=None):
    folder = os.getcwd()
    frame_path = os.path.join(folder, FRAME_NAME)
    screens_folder = os.path.join(folder, SCREENS_DIR)
    output_folder = os.path.join(folder, OUTPUT_DIR)

    if not os.path.exists(screens_folder):
        os.makedirs(screens_folder)
        log("Created: screens/ folder")

    if not os.path.exists(frame_path):
        create_placeholder_frame(frame_path)
        log("Created: frame.png (placeholder phone outline)")
        log("Tip: replace frame.png with any device mockup PNG (with transparency).")

    log(f"Frame: {frame_path}")
    log(f"Screens: {screens_folder}")
    log(f"Output: {output_folder}")

    screens = [f for f in os.listdir(screens_folder)
               if f.lower().endswith(sc.IMAGE_EXTS + sc.HEIC_EXTS)
               and not f.startswith('.')]
    if not screens:
        log("\nNo screenshots yet.")
        log(f"Drop screenshots into: {screens_folder}")
        log("Then run Phone Frame Overlay again.")
        return

    overlay_frames(frame_path, screens_folder, output_folder, log=log)


if __name__ == "__main__":
    main()
