# Design

## File layout

```
source_docs/
  *.pdf                    ← ECMA-376 source PDFs (input, not modified)
  transitional-xsd/        ← OOXML XSD schema files (input, not modified)
  strict-xsd/              ← Strict-conformance XSD (not used by default)
skills/ooxml/
  SKILL.md                 ← Claude Code skill definition
  build_index.py           ← PDF extraction + FTS indexing script
  build_schema.py          ← XSD parsing + schema parent/child indexing script
  lookup.py                ← Query script invoked by the skill
  index.db                 ← SQLite: FTS5 chunks + schema_parents + schema_children
  prefix_map.py            ← Hardcoded prefix → chapter/ML mapping
```

Two build scripts must be run to fully populate the index:
1. `build_index.py` — extracts spec text from PDFs and creates the FTS index (requires `pdftotext`)
2. `build_schema.py` — parses XSD files and populates schema parent/child tables (no extra deps)

---

## PDF text extraction

**Tool:** `pdftotext -layout <file.pdf> <output.txt>`

The `-layout` flag preserves whitespace, which helps distinguish section headings (flush-left with
a section number) from body prose and code examples.

---

## Chunking strategy

Each chunk is one spec entry — typically one numbered subsection (e.g. `17.3.2.27 rPr`).

**Section heading pattern:**
```
<section_number>  <local_name> (<human_readable_title>)
```
Examples:
- `17.3.2.27    rPr (Run Properties)`
- `19.2.1.39    sldSz (Presentation Slide Size)`
- `20.1.2.2.4   cNvCxnSpPr (Non-Visual Connector Shape Drawing Properties)`

**Parsing:** regex `^\d+(\.\d+)+\s{2,}\S` detects section boundaries. Everything from that line
to the next match is one chunk. Chunks that are pure TOC entries (no body text, just a page
number) are skipped.

**Chunk fields stored in the DB:**

| Field | Example | Notes |
|-------|---------|-------|
| `section` | `17.3.2.27` | |
| `local_name` | `rPr` | |
| `title` | `Run Properties` | |
| `ml_type` | `WordprocessingML` | derived from chapter number |
| `prefixes` | `w:` | comma-separated for shared elements |
| `source_part` | `1` | which PDF part |
| `body` | *(full section text)* | used for FTS |

Elements with the same local name in multiple MLs (e.g. `rPr` in `w:`, `a:`) are stored as
separate chunks with different `section` and `ml_type`.

Part 4 (Transitional) elements are tagged `source_part=4` and `ml_type="Transitional"` to
distinguish them from normative Part 1 entries.

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

`build_schema.py` parses XSD files from `source_docs/transitional-xsd/` and populates two
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

## Query script (`lookup.py`)

```
python lookup.py <query> [--limit N] [--part 1|2|3|4] [--summary]
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
   - Second: `local_name OR title` FTS match
   - Third: full `body` FTS match (catches descriptive queries)

**Output format:**
```
=== w:rPr - Run Properties ===
Namespace: http://schemas.openxmlformats.org/wordprocessingml/2006/main
Source: ECMA-376 Part 1,  17.3.2.27 (WordprocessingML)

This element specifies the set of properties applied to the region of text owned
by the parent run. Each property is specified as a child element of this element.
...

Parents: body, comment, tc, txBody, ...
Children:
  sequence:
    ...

Attributes:
  rsidDel  — Revision Identifier for Run Deletion
  rsidR    — Revision Identifier for Run
  ...
```

When multiple results are returned, entries are separated by a horizontal rule.

With `--summary`, only the header, namespace, and first paragraph are shown (no
parents/children/attributes). Useful for quick identification when browsing many hits.

Exit code 0 = results found; 1 = no results.
