# ooxml-lookup

A Claude skill for looking up ECMA-376 OOXML reference documentation from the source PDFs.

## Setup

1. Place the four ECMA-376 PDFs in `source_docs/` (see filenames in `build_index.py`)
2. Build the index:
   ```bash
   python skills/ooxml/build_index.py
   ```
3. Verify:
   ```bash
   python skills/ooxml/lookup.py --check
   ```

The index is stored in `skills/ooxml/index.db` (not committed — regenerate locally).
No Python dependencies beyond the standard library.

## Usage

```bash
python skills/ooxml/lookup.py "a:rPr"
python skills/ooxml/lookup.py "w:rPr"
python skills/ooxml/lookup.py "bold text"
python skills/ooxml/lookup.py "a:solidFill" --limit 1
```

## System dependencies

| Package  | Install command        | Purpose                           |
|----------|------------------------|-----------------------------------|
| poppler  | `brew install poppler` | PDF text extraction (`pdftotext`) |
