"""
Dev-only self-test for pdf_to_md.

Generates synthetic PDFs (via reportlab) covering each conversion feature,
runs the converter, and asserts the expected Markdown appears. Deterministic
round-trip check - no fixtures, no AI.
"""

import shutil
import sys
from pathlib import Path

from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle

sys.path.insert(0, str(Path(__file__).parent))
import script  # noqa: E402

WORK = Path(__file__).parent / "selftest_work"


def fit_line(seed_words, target=470):
    filler = ["alpha", "beta", "gamma", "delta", "epsilon", "zeta"]
    words = list(seed_words)
    i = 0
    while stringWidth(" ".join(words), "Helvetica", 12) < target:
        words.append(filler[i % len(filler)])
        i += 1
    return " ".join(words)


def make_features(path):
    c = canvas.Canvas(str(path), pagesize=letter)

    c.setFont("Helvetica", 24)
    c.drawString(72, 760, "Feature Test Document")

    c.setFont("Helvetica", 18)
    c.drawString(72, 720, "Chapter One")

    c.setFont("Helvetica", 12)
    c.drawString(72, 690, "Regular paragraph text with plain words.")

    c.setFont("Helvetica-Bold", 12)
    c.drawString(72, 670, "Bold intro line")

    c.setFont("Helvetica", 12)
    line1 = fit_line(["This", "sentence", "is", "deliberately", "split",
                      "across", "consecutive", "lines", "and", "ends", "with"]) + " con-"
    c.drawString(72, 650, line1)
    c.drawString(72, 636, "secutive hyphen repair as expected here.")

    c.drawString(72, 590, "\u2022 First bullet point")
    c.drawString(72, 576, "\u2022 Second bullet point")
    c.drawString(72, 562, "1. Ordered step one")
    c.drawString(72, 548, "2. Ordered step two")

    c.drawString(72, 500, "Mix of ")
    c.setFont("Helvetica-Bold", 12)
    c.drawString(120, 500, "bold")
    c.setFont("Helvetica-Oblique", 12)
    c.drawString(145, 500, "italic")
    c.setFont("Helvetica", 12)
    c.drawString(172, 500, "in one line.")

    c.setFont("Courier", 10)
    c.drawString(72, 480, "def greet(name):")
    c.drawString(72, 468, "    return 'hi ' + name")

    c.setFont("Helvetica", 12)
    c.drawString(72, 430, "Visit the documentation portal")
    c.linkURL("https://example.com/docs", (72, 420, 260, 445))

    data = [["Name", "Qty"], ["Widget", "3"], ["Gadget", "7"]]
    table = Table(data, colWidths=[120, 60])
    table.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 10),
        ("GRID", (0, 0), (-1, -1), 0.5, "black"),
    ]))
    width, height = table.wrapOn(c, 400, 200)
    table.drawOn(c, 72, 250)

    c.save()


def make_running_heads(path):
    c = canvas.Canvas(str(path), pagesize=letter)
    bodies = [
        "Overview section text for page one.",
        "Methods section text for page two.",
        "Results section text for page three.",
        "Conclusions section text for page four.",
    ]
    for i, body in enumerate(bodies, 1):
        c.setFont("Helvetica", 10)
        c.drawString(72, 760, "ACME Internal Report")
        c.setFont("Helvetica", 12)
        c.drawString(72, 650, body)
        c.setFont("Helvetica", 10)
        c.drawString(72, 30, str(i))
        c.showPage()
    c.save()


def make_scanned(path, full=True):
    png = WORK / "_scan.png"
    Image.new("RGB", (40, 30), (30, 90, 200)).save(png)

    c = canvas.Canvas(str(path), pagesize=letter)
    if not full:
        c.setFont("Helvetica", 12)
        c.drawString(72, 700, "Hello markdown world this is a text based page.")
        c.showPage()
    pages = 2 if full else 1
    for _ in range(pages):
        c.drawImage(str(png), 100, 400, 200, 150)
        c.showPage()
    c.save()


CHECKS = []


def check(name, md, expected, present=True):
    ok = (expected in md) if present else (expected not in md)
    CHECKS.append((name, ok, expected, present))


def main():
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir()

    make_features(WORK / "features.pdf")
    make_running_heads(WORK / "running_heads.pdf")
    make_scanned(WORK / "partially_scanned.pdf", full=False)
    make_scanned(WORK / "fully_scanned.pdf", full=True)

    log_lines = []
    script.convert_all_in_folder(WORK, log=log_lines.append)
    log = "\n".join(log_lines)

    if (WORK / "features.md").exists():
        md = (WORK / "features.md").read_text(encoding="utf-8")
        check("h1 heading", md, "# Feature Test Document")
        check("h2 heading", md, "## Chapter One")
        check("bold short heading", md, "#### Bold intro line")
        check("hyphen join", md, "consecutive hyphen repair")
        check("bullet list", md, "- First bullet point")
        check("bullet list 2", md, "- Second bullet point")
        check("ordered list", md, "1. Ordered step one")
        check("ordered list 2", md, "2. Ordered step two")
        check("inline bold+italic", md, "Mix of **bold** *italic* in one line.")
        check("code fence", md, "```")
        check("code content", md, "def greet(name):")
        check("link", md, "[Visit the documentation portal](https://example.com/docs)")
        check("table header", md, "| Name | Qty |")
        check("table row", md, "| Widget | 3 |")
        check("table row 2", md, "| Gadget | 7 |")
    else:
        CHECKS.append(("features.md created", False, "file missing", True))

    if (WORK / "running_heads.md").exists():
        md = (WORK / "running_heads.md").read_text(encoding="utf-8")
        check("running header removed", md, "ACME Internal Report", present=False)
        check("body page 1 kept", md, "Overview section text for page one.")
        check("body page 4 kept", md, "Conclusions section text for page four.")
        check("page numbers removed", md, "Conclusions section text for page four. 1",
              present=False)
    else:
        CHECKS.append(("running_heads.md created", False, "file missing", True))

    if (WORK / "partially_scanned.md").exists():
        md = (WORK / "partially_scanned.md").read_text(encoding="utf-8")
        check("partial scanned text kept", md, "Hello markdown world")
        check("partial scanned note", md, "page 2 skipped: scanned")
        check("partial scanned image note", md, "<!-- image: page 2, 200x150pt -->")
    else:
        CHECKS.append(("partially_scanned.md created", False, "file missing", True))

    check("fully scanned rejected", log, "appears to be scanned")
    check("fully scanned no md", "exists" if (WORK / "fully_scanned.md").exists() else "",
          "exists", present=False)

    passed = sum(1 for _, ok, _, _ in CHECKS if ok)
    failed = len(CHECKS) - passed
    print(f"\n{passed}/{len(CHECKS)} checks passed")
    for name, ok, expected, present in CHECKS:
        if not ok:
            mode = "should contain" if present else "should NOT contain"
            print(f"  FAIL: {name} ({mode}: {expected!r})")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
