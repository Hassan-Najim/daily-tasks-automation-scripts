"""
Automation Scripts - Main Menu
A unified CLI to run all automation scripts from one place.
"""

import os
import sys
import importlib.util
from pathlib import Path


# Script registry: (display_name, relative_path_to_script)
SCRIPTS = {
    "pdf": [
        ("Images to PDF", "scripts/pdf/images_to_pdf/script.py"),
        ("Images to PDF (Custom Aspect)", "scripts/pdf/images_to_pdf_custom_aspect/script.py"),
        ("Word to PDF", "scripts/pdf/word_to_pdf/script.py"),
    ],
    "image": [
        ("Phone Frame Overlay", "scripts/image/phone_frame_overlay/script.py"),
        ("Image Format Converter", "scripts/image/image_format_converter/script.py"),
    ],
    "document": [
        ("Word Find & Replace", "scripts/document/word_find_replace/script.py"),
    ],
}

CATEGORY_NAMES = {
    "pdf": "PDF",
    "image": "Image",
    "document": "Document",
}


def clear_screen():
    """Clear the terminal screen."""
    os.system('cls' if os.name == 'nt' else 'clear')


def get_script_path(relative_path: str) -> Path:
    """Get absolute path to a script."""
    base_dir = Path(__file__).resolve().parent
    return base_dir / relative_path


def run_script(script_path: Path) -> None:
    """
    Run a script by importing and calling its main() function.
    """
    if not script_path.exists():
        print(f"Error: Script not found at {script_path}")
        return

    # Change to script directory so relative paths work
    original_dir = os.getcwd()
    os.chdir(script_path.parent)

    try:
        # Load the module
        spec = importlib.util.spec_from_file_location("script", script_path)
        module = importlib.util.module_from_spec(spec)
        sys.modules["script"] = module
        spec.loader.exec_module(module)

        # Run main() if it exists
        if hasattr(module, 'main'):
            print(f"\n{'='*50}")
            print(f"Running: {script_path.parent.name}")
            print('='*50 + "\n")
            module.main()
        else:
            print(f"Error: No main() function in {script_path}")
    except Exception as e:
        print(f"Error running script: {e}")
    finally:
        os.chdir(original_dir)
        # Clean up
        if "script" in sys.modules:
            del sys.modules["script"]


def display_menu() -> None:
    """Display the main menu."""
    print("\n" + "="*50)
    print("   AUTOMATION SCRIPTS")
    print("="*50)

    index = 1
    script_map = {}

    for category, scripts in SCRIPTS.items():
        print(f"\n[{CATEGORY_NAMES[category]}]")
        for name, path in scripts:
            print(f"  {index}. {name}")
            script_map[index] = path
            index += 1

    print(f"\n  q. Quit")
    print("="*50)

    return script_map


def main():
    """Main menu loop."""
    while True:
        script_map = display_menu()

        choice = input("\nEnter choice: ").strip().lower()

        if choice == 'q':
            print("\nGoodbye!")
            break

        try:
            choice_num = int(choice)
            if choice_num in script_map:
                script_path = get_script_path(script_map[choice_num])
                run_script(script_path)
                input("\nPress Enter to continue...")
            else:
                print("Invalid choice. Please try again.")
        except ValueError:
            print("Invalid input. Enter a number or 'q' to quit.")


if __name__ == "__main__":
    main()
