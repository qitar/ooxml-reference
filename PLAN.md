# OOXML Reference Lookup Skill — Implementation Plan

## Overview

A Claude Code skill that enables looking up ECMA-376 OOXML specification entries directly
from the source PDFs. Designed to supplement pptx/xlsx editing and diff-inspection skills
by answering questions like "what does `a:rPr` do?" or "how do I set bold text in DrawingML?".

---

## Source Material

| File | Pages | Content |
|------|-------|---------|
| Part 1 — Fundamentals & Markup Language Reference | 5039 | Main reference: WordprocessingML (ch17), SpreadsheetML (ch18), PresentationML (ch19), DrawingML (ch20–21), Shared (ch22) |
| Part 2 — Open Packaging Conventions | 95 | Package/part/relationship structure |
| Part 3 — Markup Compatibility & Extensibility | 43 | mc: namespace, fallback handling |
| Part 4 — Transitional Migration Features | 1553 | Legacy/transitional elements |

Part 1 is the primary reference for day-to-day OOXML work. Parts 2–4 are niche but worth indexing.

---

## Namespace Prefix Mapping

Section headings in the spec use **local names only** (e.g. `rPr`, `sldSz`). The chapter
structure provides the namespace context. This mapping is hardcoded and stable:

| Prefix | Namespace URI | Chapters (Part 1) | ML Type |
|--------|---------------|-------------------|---------|
| `w:`   | `.../wordprocessingml/main` | 17 | WordprocessingML |
| `x:`   | `.../spreadsheetml/main` | 18 | SpreadsheetML |
| `p:`   | `.../presentationml/main` | 19 | PresentationML |
| `a:`   | `.../drawingml/main` | 20–21 (main) | DrawingML Main |
| `c:`   | `.../drawingml/chart` | 21.2–21.3 | DrawingML Charts |
| `dgm:` | `.../drawingml/diagram` | 21.4 | DrawingML Diagrams |
| `r:`   | `.../officeDocument/relationships` | 22.8 | Relationships |
| `mc:`  | `.../markup-compatibility` | Part 3 | Markup Compatibility |

Note: `a:rPr` (DrawingML run properties) and `w:rPr` (WordprocessingML run properties)
are distinct elements with different attributes. Prefix-aware routing is essential.

---

## Architecture

```
source_docs/         ← PDFs (input, not modified)
skills/ooxml/
  SKILL.md           ← Claude Code skill definition
  build_index.py     ← One-time extraction + indexing script
  lookup.py          ← Query script invoked by the skill
  index.db           ← SQLite FTS5 index
  prefix_map.py      ← Hardcoded prefix → chapter/ML mapping
```

The index is built once (or rebuilt when PDFs change) and queried at skill invocation time.

---

## Step 1: PDF Text Extraction

**Tool:** `pdftotext` (poppler, installed to the system)

**Command:** `pdftotext -layout <file.pdf> <output.txt>`

The `-layout` flag preserves whitespace formatting, which helps distinguish section headings
(which are flush-left with a section number) from body prose and code examples.

**Output:** One `.txt` file per PDF, stored temporarily during the build step.

---

## Step 2: Chunking Strategy

Each "chunk" is one spec entry — typically one numbered subsection (e.g. `17.3.2.27 rPr`).

**Section heading pattern:**
```
<section_number>  <local_name> (<human_readable_title>)
```
Examples:
- `17.3.2.27    rPr (Run Properties)`
- `19.2.1.39    sldSz (Presentation Slide Size)`
- `20.1.2.2.4   cNvCxnSpPr (Non-Visual Connector Shape Drawing Properties)`

**Parsing approach:**
- Regex to detect lines matching `^\d+(\.\d+)+\s{2,}\S` (section number + 2+ spaces + content)
- Everything from that line until the next matching line = one chunk
- Chunks that are pure TOC entries (no body text, just a page number) are skipped

**Chunk fields stored in the DB:**
- `section` — e.g. `17.3.2.27`
- `local_name` — e.g. `rPr`
- `title` — e.g. `Run Properties`
- `ml_type` — e.g. `WordprocessingML` (derived from chapter number)
- `prefixes` — e.g. `w:` (derived from chapter; comma-separated for shared elements)
- `source_part` — e.g. `1` (which PDF part)
- `body` — full text of the section for FTS

---

## Step 3: SQLite FTS5 Index

**Schema:**

```sql
-- Main content store
CREATE TABLE chunks (
    id          INTEGER PRIMARY KEY,
    section     TEXT,       -- "17.3.2.27"
    local_name  TEXT,       -- "rPr"
    title       TEXT,       -- "Run Properties"
    ml_type     TEXT,       -- "WordprocessingML"
    prefixes    TEXT,       -- "w:"
    source_part INTEGER,    -- 1, 2, 3, or 4
    body        TEXT        -- full section text
);

-- FTS5 virtual table, content= to avoid duplicating body text
CREATE VIRTUAL TABLE chunks_fts USING fts5(
    local_name,
    title,
    body,
    content='chunks',
    content_rowid='id'
);
```

**Why FTS5 over FTS4:** FTS5 supports `rank` for relevance ordering, phrase queries, and
`highlight()`/`snippet()` for result excerpts — all useful for this skill.

---

## Step 4: Query Script (`lookup.py`)

**Usage:**
```
python lookup.py <query> [--limit 5] [--part 1|2|3|4] [--check]
```

**Query resolution logic:**

1. If query contains a `:` (e.g. `w:rPr`, `a:solidFill`):
   - Strip prefix → `rPr`
   - Map prefix → ML type filter (e.g. `w:` → `WordprocessingML`)
   - Run FTS on `local_name` with ML type filter, then fall back to broader FTS if no results

2. If query has no prefix (e.g. `rPr`, `bold text run`):
   - Run FTS across all chunks
   - Group/rank results by ML type so the user can see all contexts

3. FTS query construction:
   - First try: exact match on `local_name` (highest precision)
   - Second try: `local_name OR title` FTS match
   - Third try: full `body` FTS match (catches descriptive queries)

**Output format (plain text, for Claude to read):**

```
=== w:rPr — Run Properties (17.3.2.27) [WordprocessingML] ===

This element specifies the set of properties applied to the region of text owned
by the parent run. Each property is specified as a child element of this element.
...

Attributes:
  rsidDel  — Revision Identifier for Run Deletion
  rsidR    — Revision Identifier for Run
  ...

[Snippet from body if result is from FTS body match, not exact local_name match]
```

**Exit codes:** 0 = results found, 1 = no results (so Claude can report clearly).

---

## Step 5: SKILL.md

The skill definition tells Claude:

1. **When to activate** — when inspecting unknown XML nodes/attributes in pptx/xlsx diffs,
   or when needing to know how to express a visual property in OOXML.

2. **How to invoke** — run `python skills/ooxml/lookup.py "<query>"` from the project root,
   or with an absolute path to the script.

3. **How to interpret results** — section number, ML type, attributes table, prose description.

4. **Prefix stripping hint** — remind Claude that `w:rPr` → query `w:rPr` directly (the script
   handles stripping); or query just `rPr` to see all ML contexts.

5. **Fallback** — if no results, suggest trying the local name without prefix, or a descriptive
   phrase.

---

## Step 6: Build Script (`build_index.py`)

Orchestrates the full pipeline:

1. Extract text from all 4 PDFs via `pdftotext`
2. Parse and chunk each extracted text file
3. Create the SQLite DB and FTS5 index
4. Print a summary: chunks indexed per part, per ML type

**Runtime estimate:** ~2–5 minutes for all ~6730 pages (dominated by `pdftotext` on Part 1).

**Rebuild trigger:** Run manually when PDFs change. A `--force` flag drops and recreates the DB.

---

## Open Questions / Risks

1. **`pdftotext` output quality** — The PDFs were generated by Word 2013. Column layout and
   tables (attribute tables especially) may extract poorly. The attribute table rows in the
   spec are critical content; we may need post-processing to flatten them into readable text
   rather than misaligned columns.

2. **TOC pages** — Part 1 has a very long TOC (pages 1–186 roughly). The chunker needs to
   detect and skip TOC-only entries that lack body text, or they'll pollute results with
   duplicate shallow matches.

3. **Same local name, multiple MLs** — e.g. `rPr` in `w:`, `a:`, and others. The index stores
   them as separate chunks (different `section`, `ml_type`). The query script returns all
   of them when no prefix filter is given, ordered by ML type.

4. **Sub-section depth** — Some elements have many sub-subsections for each attribute.
   Decide at build time whether to chunk at the element level (one chunk per element,
   attributes included in body) or at the attribute level (one chunk per attribute). Element-
   level chunking is simpler and sufficient for the primary use case.

5. **Part 4 (Transitional)** — Contains many elements with the same names as Part 1 but
   with legacy semantics. These are tagged with `source_part=4` and `ml_type="Transitional"` so
   results distinguish them from the normative Part 1 entries.

---

## Implementation Order

All steps completed in this order:

1. `build_index.py` — extraction + chunking + indexing (validates the whole pipeline)
2. Manual inspection of a sample of indexed chunks to verify quality
3. `lookup.py` — query script
4. `prefix_map.py` — namespace mapping table
5. `SKILL.md` — skill definition (written last, once we know exactly how the script works)
