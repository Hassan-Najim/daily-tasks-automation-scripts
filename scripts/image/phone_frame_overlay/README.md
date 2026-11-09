# Phone Frame Overlay

## WHAT

Places your screenshots inside a phone frame image for polished mockups.

## NEEDS

- Screenshots (PNG, JPG, BMP, GIF, TIFF, WEBP, HEIC) in a `screens/` subfolder
- `frame.png` device mockup (with transparency) - both are created
  automatically on first run

## HOW

1. Run the script once - it creates `screens/` and a placeholder `frame.png`:

```bash
python script.py
```

2. Drop your screenshots into `screens/`
3. Run it again

## OUTPUT

- `output/` folder with a framed copy of each screenshot

## Tips

- Replace the placeholder `frame.png` with any device mockup
  (transparent area where the screen shows)
- Screenshots larger than the frame are scaled down to fit
- Framed copies are always saved as PNG

## Folder Structure

```
phone_frame_overlay/
├── script.py
├── frame.png          <- Your device frame (auto-created as placeholder)
├── screens/           <- Your screenshots go here (auto-created)
│   ├── screen1.png
│   └── screen2.png
└── output/            <- Created automatically
    ├── screen1.png
    └── screen2.png
```
