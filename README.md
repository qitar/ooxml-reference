# ooxml-reference

An AI agent skill for looking up the ECMA-376 (ISO/IEC 29500) OOXML reference.
Uses an SQLite index built from ECMA-376 specification PDFs and XSDs
and a lookup script that the agent can query using element names or description text.

## Source material

ECMA-376 5th Edition (2016).

| Part | Pages | Content |
|------|-------|---------|
| Part 1 — Fundamentals & Markup Language Reference | 5039 | Main reference |
| Part 2 — Open Packaging Conventions | 95 | Package/part/relationship structure |
| Part 3 — Markup Compatibility & Extensibility | 43 | `mc:` (Markup Compatibility) namespace, fallback handling |
| Part 4 — Transitional Migration Features | 1553 | Legacy/transitional elements |

## Building index.db

From the `skills/ooxml-reference/scripts/` directory:

```bash
python build_index.py
python build_schema.py
```

The index is stored in `skills/ooxml-reference/scripts/index.db`.

## Build-time dependencies

| Package  | Install command        | Purpose                           |
|----------|------------------------|-----------------------------------|
| poppler  | `brew install poppler` | PDF text extraction (`pdftotext`) |
