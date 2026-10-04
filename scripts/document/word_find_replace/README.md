# Word Document Find & Replace

## WHAT

Applies bulk sentence-level replacements to a Word document from a corrections list.

## NEEDS

- A .docx document (path set in `script.py`)
- Corrections defined inline in `script.py` or loaded from a JSON file
- Python packages: python-docx

## HOW

1. Edit `main()` in `script.py` - set `doc_path` to your document and
   define your corrections inline (or load them from a JSON file):

```python
corrections = [
    {
        "OgSentence": "Original text to find",
        "NewSentence": "Replacement text"
    }
]

doc_path = "document.docx"  # Change this to your document path
```

2. Run the script:

```bash
python script.py
```

## OUTPUT

- A copy of the document with corrections applied, saved as `*_updated.docx`
- The original document is never modified

## Notes

- Searches paragraphs and tables
- Replacements are case-sensitive and exact-match
- Only the first occurrence of each sentence is replaced
