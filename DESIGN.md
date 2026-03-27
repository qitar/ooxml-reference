# Design

## File layout

```
skills/ooxml-reference/
  SKILL.md                 ← Claude Code skill definition
  pdfs/                    ← ECMA-376 5th Edition (2016) source PDFs
  schemas/                 ← OOXML transitional XSD schema files
  scripts/
    lookup.py              ← Thin entry point invoked by the skill
    _query.py              ← Query logic (two-stage FTS fallback)
    build.py               ← Entry point to rebuild the full index
    _build_index.py        ← PDF extraction + FTS indexing
    _build_schema.py       ← XSD parsing + schema parent/child indexing
    _prefix_map.py         ← Hardcoded prefix → chapter/ML mapping
    index.db               ← SQLite: FTS5 chunks + schema_parents + schema_children
```

## PDF text extraction

**Tool:** pymupdf (`fitz`) — extracts text with per-span font metadata (name, size, color).

Each page is processed via `page.get_text("dict")`, which returns structured blocks/lines/spans.
Font metadata on each span distinguishes headings from body text without regex heuristics.

**Dependency:** `PyMuPDF` (imported as `fitz`). Not installed globally; invoked via
`uv run --with pymupdf python build.py`.

---

## Chunking strategy

Each chunk is one spec entry — typically one numbered subsection (e.g. `17.3.2.27 rPr`).

**Section heading detection:** headings are identified by font metadata.
A text block is a heading block when any line's first span uses font `Cambria*` at
size >= 12pt. Body text uses `Calibri 11pt` and code examples use `Consolas 10-11pt`,
so neither triggers heading detection.

Once a heading block is found, all its lines are joined and parsed with a simple regex
to extract the section number, element name, and parenthesized title:
```
17.3.2.27 rPr (Run Properties)
19.2.1.39 sldSz (Presentation Slide Size)
20.1.2.2.4 cNvCxnSpPr (Non-Visual Connector Shape Drawing Properties)
```

Heading-styled blocks without a dotted section number (e.g. "Foreword", chapter titles)
are treated as body text. Page headers (y < 70pt) and footers (y > 730pt) are skipped.

Chunks whose `section_to_ml` resolves to `"Unknown"` are dropped, as are chunks with
fewer than 50 characters of body text.

**Body text assembly:** pymupdf often splits side-by-side table cells (e.g. an
attribute name column and its description column) into separate blocks that share
the same y-position. To preserve column layout, consecutive non-heading blocks on
each page are batched and their lines are merged by vertical position: lines whose
midpoints are within half a line-height are grouped into a single row. Within each
row, spans are sorted by x-position and joined with proportional spacing when the
horizontal gap exceeds ~3 character widths (`_spans_to_text`).

**Attribute table formatting:** the "Attributes / Description" tables get special
treatment. When `merge_block_lines` detects the header row (via `TABLE_HEADER_RE`),
it enters a three-phase table mode:

1. **Header:** emit a fixed-width header and set a sentinel value for `table_col_x`.
2. **First body row:** find the largest inter-span gap (>20pt) and use its midpoint
   as the column boundary. The header text position is unreliable for this because
   "Description" can sit far to the right of where body descriptions actually start.
3. **Subsequent rows:** split each row's spans into left/right by the boundary,
   format each side with `_spans_to_text`, and pad the left column to a fixed
   width (`ATTR_COL_WIDTH`). Rows with only left-column text shorter than
   `ATTR_COL_WIDTH` are attribute-name continuations; rows with only long
   left-column text signal the table has ended (e.g. the `[Note: ...` paragraph).

`table_col_x` state is threaded through `flush_body_blocks` across page breaks
so multi-page tables format consistently. It resets to `None` at each new chunk.

**Chunk fields stored in the DB:**

| Field | Example | Notes |
|-------|---------|-------|
| `section` | `17.3.2.27` | |
| `local_name` | `rPr` | NULL for section-level headings with no parenthesized name (e.g. `17.1 WordprocessingML`) |
| `title` | `Run Properties` | |
| `ml_type` | `WordprocessingML` | derived from section number via `CHAPTER_MAP` / `SUBSECTION_MAP` in `_prefix_map.py` |
| `prefixes` | `w:` | comma-separated for shared elements |
| `source_part` | `1` | which PDF part |
| `body` | *(full section text)* | used for FTS |

Elements with the same local name in multiple MLs (e.g. `rPr` in `w:`, `a:`) are stored as
separate chunks with different `section` and `ml_type`.

Part 4 (Transitional) elements are tagged `source_part=4` and `ml_type="Transitional"` to
distinguish them from normative Part 1 entries.

---

## Namespace / prefix mapping

`_prefix_map.py` contains three declarative structures that drive namespace resolution
across both the PDF chunker and the XSD schema builder:

- **`PREFIX_MAP`** maps namespace prefix → (ml_type, namespace URI). This is the single
  source of truth for all known prefixes. Derived maps (`ML_TO_PREFIX`, `URI_TO_PREFIX`,
  `URI_TO_ML`) are computed automatically from it.

- **`CHAPTER_MAP`** maps Part 1 top-level chapter numbers to a default (ml_type, prefix).

- **`SUBSECTION_MAP`** overrides `CHAPTER_MAP` for chapters that contain multiple
  sub-namespaces. Keyed by `(chapter, subsection)` tuples, e.g. `(20, 4)` for section
  20.4 (WordprocessingDrawing). The lookup function `section_to_ml` checks `SUBSECTION_MAP`
  first, then falls back to `CHAPTER_MAP`.

DrawingML sub-namespaces (`wp`, `xdr`, `cdr`, `pic`, `lc`) each have their own
`PREFIX_MAP` entry so that prefixed queries like `wp:anchor` resolve correctly and
schema tables use the right ml_type.

### The `pic` edge case

The `pic:` namespace (`dml-picture.xsd`, 6 elements) has no dedicated chapter section
in the spec PDF. Its elements are documented inline within sections 20.1, 20.2, and 20.5,
so their chunks get tagged with the containing section's ml_type (`DrawingML`,
`DrawingML Chart Drawing`, etc.) rather than `DrawingML Picture`. Meanwhile, the schema
tables derive ml_type from the XSD's target namespace URI, which maps to
`DrawingML Picture`. This means a prefixed query like `pic:blipFill` returns no results
because both query stages filter on ml_type and no chunks carry `DrawingML Picture`.
Unprefixed queries (e.g. just `blipFill`) work fine and return all contexts.

---

## SQLite FTS5 index

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

-- content= avoids duplicating body text on disk
CREATE VIRTUAL TABLE chunks_fts USING fts5(
    local_name,
    title,
    body,
    content='chunks',
    content_rowid='id'
);
```

FTS5 is used over FTS4 because it supports `rank` for relevance ordering, phrase queries, and
`highlight()`/`snippet()` for result excerpts.

---

## Schema index

`_build_schema.py` parses XSD files from `schemas/` and populates two
additional tables in `index.db`:

```sql
-- Comma-separated list of element local names that may contain this element,
-- derived by inverting each complex type's content model.
CREATE TABLE schema_parents (local_name TEXT, ml_type TEXT, parents TEXT);

-- Pre-rendered indented text showing the element's content model
-- (sequence/choice/all structure with cardinalities).
CREATE TABLE schema_children (local_name TEXT, ml_type TEXT, content_model TEXT);
```

**Why transitional XSD:** namespace URIs match `PREFIX_MAP`; transitional is a superset of strict,
so it gives the broadest coverage.

**Implementation notes:**
- Element → CT mappings are collected from all `<xsd:element name="X" type="CT_X">` declarations
  throughout each file, not just global declarations. OOXML XSD uses local element declarations
  inside complex types extensively, so walking the full tree is required.
- `xsd:group` definitions are expanded inline during rendering up to depth 8.
- `minOccurs`/`maxOccurs` are rendered as `[min..*]`; absent means required exactly once.
- The script is idempotent: drops and recreates its two tables on each run.
- `lookup.py` silently omits schema sections if the tables don't exist yet.

**Output format in `lookup.py`:**
```
Parents: body, comment, tc, txBody, ...
Children:
  sequence:
    pPr [0..1]
    choice [0..*]:
      r
      hyperlink
      ...
```

---

## Query script (`scripts/lookup.py`)

```
python scripts/lookup.py <query> [--limit N] [--part 1|2|3|4] [--summary]
```

**Query resolution:**

1. Query contains `:` (e.g. `w:rPr`, `a:solidFill`):
   - Strip prefix → `rPr`
   - Map prefix → ML type filter (e.g. `w:` → `WordprocessingML`)
   - Run FTS on `local_name` with ML type filter; fall back to broader FTS if no results

2. Query has no prefix (e.g. `rPr`, `bold text run`):
   - Run FTS across all chunks
   - Group/rank results by ML type so all contexts are visible

3. FTS cascade:
   - First: exact match on `local_name` (highest precision)
   - Second: tokenized FTS (query split into words, implicit AND) with `bm25` column weights (10, 5, 1) so `local_name`/`title` hits rank above body-only hits

**Output format:**
```
========================================================================
w:rPr - Run Properties (WordprocessingML)
========================================================================

This element specifies the set of properties applied to the region of text owned
by the parent run. Each property is specified as a child element of this element.
...

Parents:
body, comment, tc, txBody, ...

Children:
  sequence:
    ...

Namespace: http://schemas.openxmlformats.org/wordprocessingml/2006/main
Source: ECMA-376 Part 1, 17.3.2.27
```

When multiple results are returned, entries are separated by blank lines.

With `--summary`, only the header, namespace, and first paragraph are shown (no
parents/children/attributes). Useful for quick identification when browsing many hits.

Exit code 0 = results found; 1 = no results.
