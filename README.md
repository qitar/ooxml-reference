# ooxml-reference

A CLI script + an AI agent skill for looking up the ECMA-376 (ISO/IEC 29500) OOXML reference.

OOXML is the XML-based file format behind PowerPoint (.pptx),
Word (.docx), and Excel (.xlsx) files.

Uses an SQLite index built from ECMA-376 5th Edition (2016) PDFs and XSDs
and a lookup script that agents can query using element names or description text.

## Script usage

```
python scripts/lookup.py "a:rPr"
python scripts/lookup.py "bold text"
```

See skills/ooxml-reference/SKILL.md for details.

## Skill usage

```
npx skills add https://github.com/qitar/ooxml-reference --skill ooxml-reference

claude "How do I change table cell background color to a gradient with OOXML?"
```

See skills/ooxml-reference/SKILL.md for details.

## Updating source PDFs and XSDs

Grab the latest version at https://ecma-international.org/publications-and-standards/standards/ecma-376/ .

Transitional XSD Schema files are included in the Part 4 zip file.

The build script will likely have to be revised to accomodate formatting changes in the PDFs.

## Building index.db

From the `skills/ooxml-reference/scripts/` directory:

```bash
uv run --with pymupdf python build.py
```

The index is stored in `skills/ooxml-reference/scripts/index.db`.

## Build-time dependencies

| Package  | Install command                        | Purpose                              |
|----------|----------------------------------------|--------------------------------------|
| PyMuPDF  | `uv run --with pymupdf` or `pip install pymupdf` | PDF parsing with font metadata |
