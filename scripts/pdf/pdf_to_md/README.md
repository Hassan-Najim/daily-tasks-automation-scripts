# PDF to Markdown

## WHAT

Batch-converts text-based PDFs to Markdown using deterministic layout
heuristics — no AI, no OCR.

## NEEDS

- Text-based PDFs in this folder (scanned pages are detected and skipped)
- Python packages: pdfplumber (`pip install -r requirements.txt`)

## HOW

1. Place your PDFs in this folder
2. Run the script:

```bash
python script.py
```

## OUTPUT

- One sibling `<name>.md` per PDF; existing `.md` files are never overwritten
- Detects headings, bold/italic/inline code, lists, tables, links and code
  blocks; strips running headers/footers and page numbers
- Scanned pages are reported and skipped (no OCR); images become
  `<!-- image: page N, WxH -->` placeholders

## Limitations

- Multi-column layouts (research papers) are not split into columns yet

## What it detects

- Headings (`#`–`####`) via font-size clustering and short bold standalone lines
- Bold, italic, and inline code via font names
- Bullet and numbered lists, including nesting levels and wrapped lines
- Tables (rendered as GitHub pipe tables, multiline cells joined with `<br>`)
- Hyperlinks rendered as `[text](url)`
- Monospace runs rendered as fenced code blocks
- Running headers/footers and page numbers (removed when repeated across pages)
- Ligature repair (`ﬁ` → `fi`) and hyphenated line-break joins

## Self-test

```
python selftest.py
```

Generates synthetic PDFs covering every feature and verifies the Markdown
round-trips correctly (24 deterministic checks).
