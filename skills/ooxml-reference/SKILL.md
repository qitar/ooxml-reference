# OOXML Reference Lookup Skill

## When to use this skill

Use this skill whenever you need authoritative ECMA-376 spec information:

- You encounter an unknown XML element or attribute in a pptx/xlsx diff (e.g. `<a:solidFill>`, `w:rPr`, `p:sldSz`)
- You need to know which element or attribute produces a specific visual effect (e.g. "how do I set bold text in DrawingML?")
- The user asks what an OOXML XML node or attribute means or does
- You are editing raw OOXML and need to know valid attributes, child elements, or constraints

## Running the lookup script

Determine the absolute path to this skill's directory (the directory containing this SKILL.md file), then invoke the lookup script relative to it:

```bash
python <skill_dir>/scripts/lookup.py "<query>"
python <skill_dir>/scripts/lookup.py "<query>" --limit 3
python <skill_dir>/scripts/lookup.py "<query>" --summary   # header + first paragraph only
```

Query forms:

| Query | Behavior |
|---|---|
| `a:solidFill` | Looks up `solidFill` filtered to DrawingML |
| `w:rPr` | Looks up `rPr` filtered to WordprocessingML |
| `rPr` | Looks up `rPr` across all ML types |
| `"bold text"` | Full-text phrase search |

Namespace prefix → ML type mapping:

| Prefix | ML type | Spec chapters |
|---|---|---|
| `w:` | WordprocessingML | ch.17 |
| `x:` | SpreadsheetML | ch.18 |
| `p:` | PresentationML | ch.19 |
| `a:` | DrawingML | ch.20 - 21 main |
| `c:` | DrawingML Charts | ch.21.2 - 21.3 |
| `dgm:` | DrawingML Diagrams | ch.21.4 |
| `r:` | Relationships | ch.22.8 |
| `m:` | Math | ch.22.1 |
| `mc:` | Markup Compatibility | Part 3 |
| `wps:` | WordprocessingML Shapes | Word 2010 extension |
| `wpg:` | WordprocessingML Group | Word 2010 extension |

Namespace prefixes in actual pptx/xlsx XML files are aliases defined per-file in `xmlns` declarations. They conventionally follow this mapping, but if you see an unusual prefix, check the `xmlns` declaration in the file to confirm the actual namespace URI, then use the conventional prefix for lookup.

### How querying works

The lookup script uses a three-stage fallback — each stage only runs if the previous one returned no results:

1. **Exact name match.** Looks for an exact `local_name` in the index (with optional `ml_type` filter from the namespace prefix). This is the fast path for element/attribute lookups like `a:solidFill` or `rPr`.
2. **FTS on name and title.** Full-text search scoped to the `local_name` and `title` fields. Catches partial matches and queries that correspond to a section title rather than an exact element name.
3. **Full-body FTS.** Searches the entire spec body text and returns a short snippet around the match. This handles descriptive queries like `"bold text"` that don't match any element name or title.

Prefer prefixed element names (stage 1) for precise results. Use natural-language phrases when you don't know the element name — they will reach stage 3 automatically.

## Included source documents

The ECMA-376 5th Edition (2016) specification PDFs and transitional XSD schema files are included in this skill under `pdfs/` and `schemas/` respectively. If the lookup index does not contain enough detail for your task, you may read the PDF or XSD files directly for additional context.

## Interpreting results

Each result block has this structure:

```
=== a:rPr - Text Run Properties ===
Namespace: http://schemas.openxmlformats.org/drawingml/2006/main
Source: ECMA-376 Part 1,  21.1.2.3.9 (DrawingML)

<spec text: prose description, attributes table, XML examples, cross-references>

Parents:
spPr, grpSpPr, ln, ...

Children:
  sequence:
    ...
```

When multiple results are returned, entries are separated by a horizontal rule (`---`).

Key fields:
- The header gives the qualified name and human title. Section reference and ML type are on the Source line.
- The spec text contains the attributes table (columns may appear flattened due to PDF extraction), usage examples, and cross-references to related sections (e.g. `20.1.10.40`).
- When a body is long, a `[Match found in body]` snippet is shown instead of the full text. If you need more detail, re-run with `--limit 1` targeting the specific element.
- **Parents** lists the element names that may contain this element, derived from the XSD schema.
- **Children** shows the content model (sequence/choice structure with cardinalities) from the XSD schema. Cardinality is shown as `[min..max]`; absent means required exactly once. `*` means `unbounded`. `...` means the nesting was truncated at depth 8.

## Caveats

**Multiple results for the same local name.** Elements like `rPr` exist in both WordprocessingML and DrawingML with different attributes. Always check the ML type on the `Source:` line to use the right spec entry for the file you are working with.

**Prefer Part 1 results.** Part 4 contains transitional/legacy elements from the original OOXML spec. For current pptx/xlsx work, prefer `Source: ECMA-376 Part 1` results. Use `--part 1` to filter explicitly if Part 4 entries are appearing.

**Missing index.** If the index has not been built yet, the script will exit with an error. To build it, `cd` into the `scripts/` directory and run `python build_index.py` (requires the PDFs in `pdfs/` and `pdftotext` from poppler), then `python build_schema.py` to populate the Parents/Children schema data.
