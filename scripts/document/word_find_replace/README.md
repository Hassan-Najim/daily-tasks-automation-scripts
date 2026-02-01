# Word Document Find & Replace

Performs bulk find-and-replace operations in Word documents using JSON-formatted corrections.

## Features

- Batch find-and-replace from JSON
- Searches paragraphs and tables
- Creates a new file (preserves original)
- Supports loading corrections from file or inline

## Usage

### 1. Prepare your corrections

Create a JSON file (`corrections.json`):

```json
[
    {
        "OgSentence": "Original text to find",
        "NewSentence": "Replacement text"
    },
    {
        "OgSentence": "Another original",
        "NewSentence": "Another replacement"
    }
]
```

### 2. Edit the script

Update `main()` with your document path and corrections source:

```python
corrections = load_corrections_from_json("corrections.json")
doc_path = "my_document.docx"
```

### 3. Run

```bash
python script.py
```

### 4. Output

Creates `my_document_updated.docx` with all replacements applied.

## JSON Format

```json
{
    "OgSentence": "The exact text to find",
    "NewSentence": "The text to replace it with"
}
```

## Notes

- Replacements are case-sensitive and exact-match
- Only the first occurrence of each sentence is replaced
- Original document is never modified
