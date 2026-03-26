# ooxml-lookup

A Claude Code skill for looking up ECMA-376 OOXML specification entries directly from the source
PDFs. Designed to supplement pptx/xlsx editing and diff-inspection skills by answering questions
like "what does `a:rPr` do?" or "how do I set bold text in DrawingML?".

## Source material

| Part | Pages | Content |
|------|-------|---------|
| Part 1 — Fundamentals & Markup Language Reference | 5039 | Main reference: WordprocessingML (ch17), SpreadsheetML (ch18), PresentationML (ch19), DrawingML (ch20–21), Shared (ch22) |
| Part 2 — Open Packaging Conventions | 95 | Package/part/relationship structure |
| Part 3 — Markup Compatibility & Extensibility | 43 | `mc:` namespace, fallback handling |
| Part 4 — Transitional Migration Features | 1553 | Legacy/transitional elements |

Part 1 is the primary reference for day-to-day OOXML work.

## Namespace prefixes

Section headings in the spec use local names only (e.g. `rPr`, `sldSz`). The chapter structure
provides the namespace context:

| Prefix | Namespace URI | Chapters (Part 1) | ML type |
|--------|---------------|-------------------|---------|
| `w:`   | `.../wordprocessingml/main` | 17 | WordprocessingML |
| `x:`   | `.../spreadsheetml/main` | 18 | SpreadsheetML |
| `p:`   | `.../presentationml/main` | 19 | PresentationML |
| `a:`   | `.../drawingml/main` | 20–21 (main) | DrawingML Main |
| `c:`   | `.../drawingml/chart` | 21.2–21.3 | DrawingML Charts |
| `dgm:` | `.../drawingml/diagram` | 21.4 | DrawingML Diagrams |
| `r:`   | `.../officeDocument/relationships` | 22.8 | Relationships |
| `mc:`  | `.../markup-compatibility` | Part 3 | Markup Compatibility |

Note: `a:rPr` (DrawingML run properties) and `w:rPr` (WordprocessingML run properties) are
distinct elements with different attributes — prefix-aware routing is essential.

## Setup

1. Place the four ECMA-376 PDFs in `source_docs/` (see filenames in `build_index.py`)
2. Build the FTS index from PDFs:
   ```bash
   python skills/ooxml/build_index.py
   ```
3. Build the schema parent/child index from XSD files (already present in `source_docs/transitional-xsd/`):
   ```bash
   python skills/ooxml/build_schema.py
   ```
4. Verify:
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
python skills/ooxml/lookup.py "rPr" --summary      # header + first paragraph only
```

## System dependencies

| Package  | Install command        | Purpose                           |
|----------|------------------------|-----------------------------------|
| poppler  | `brew install poppler` | PDF text extraction (`pdftotext`) |
