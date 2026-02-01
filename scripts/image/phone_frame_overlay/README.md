# Phone Frame Overlay

Composites screenshots onto a device frame mockup for professional-looking app previews.

## Features

- Automatically centers screenshots in the frame
- Preserves transparency
- Batch processes all PNG files in the screens folder
- Outputs high-quality PNG files

## Setup

1. Add a `frame.png` file to this folder (device mockup with transparent screen area)
2. Create a `screens/` subfolder
3. Place your screenshot PNGs in `screens/`

## Folder Structure

```
phone_frame_overlay/
├── script.py
├── frame.png          <- Your device frame
├── screens/           <- Your screenshots go here
│   ├── screen1.png
│   └── screen2.png
└── output/            <- Created automatically
    ├── screen1.png
    └── screen2.png
```

## Usage

```bash
python script.py
```

## Tips

- Frame should have a transparent area where the screen will show
- Screenshots should be sized to fit within the frame's screen area
- Use PNG format for best quality (preserves transparency)
