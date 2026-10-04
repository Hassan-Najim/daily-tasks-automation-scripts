# PDF to Markdown

Batch-converts text-based PDFs to Markdown using deterministic layout
heuristics — no AI, no OCR.

Every `.pdf` in the folder becomes a sibling `<name>.md`. Existing `.md`
files are never overwritten.

## What it detects

- Headings (`#`–`####`) via font-size clustering and short bold standalone lines
- Bold, italic, and inline code via font names
- Bullet and numbered lists, including nesting levels and wrapped lines
- Tables (rendered as GitHub pipe tables, multiline cells joined with `<br>`)
- Hyperlinks rendered as `[text](url)`
- Monospace runs rendered as fenced code blocks
- Running headers/footers and page numbers (removed when repeated across pages)
- Ligature repair (`ﬁ` → `fi`) and hyphenated line-break joins

## Limitations

- **Scanned PDFs are not supported** — pages without a text layer are
  detected, reported as a warning, and skipped (OCR is out of scope)
- Multi-column layouts (research papers) are not split into columns yet
- Images are not extracted; they become `<!-- image: page N, WxH -->` placeholders

## Requirements

```
pip install -r requirements.txt
```

## Self-test

```
python selftest.py
```

Generates synthetic PDFs covering every feature and verifies the Markdown
round-trips correctly (24 deterministic checks).
