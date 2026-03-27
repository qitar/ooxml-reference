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
python <skill_dir>/scripts/lookup.py "<query>" --part 1    # prefer Part 1 (normative) over Part 4 (transitional)
python <skill_dir>/scripts/lookup.py "<query>" --summary   # header + first paragraph only
```

Query forms:

| Query | Behavior |
|---|---|
| `a:solidFill` | Looks up `solidFill` filtered to DrawingML |
| `w:rPr` | Looks up `rPr` filtered to WordprocessingML |
| `rPr` | Looks up `rPr` across all ML types |
| `bold text` | Full-text search (all words must match, any order) |

Namespace prefix → ML type mapping:

| Prefix | ML type | Spec chapters |
|---|---|---|
| `w:` | WordprocessingML | ch.17 |
| `x:` | SpreadsheetML | ch.18 |
| `p:` | PresentationML | ch.19 |
| `a:` | DrawingML | ch.20 - 21 main |
| `c:` | DrawingML Charts | ch.21.2 - 21.3 |
| `dgm:` | DrawingML Diagrams | ch.21.4 |
| `m:` | Math | ch.22.1 |
| `r:` | Relationships | ch.22.8 |
| `mc:` | Markup Compatibility | Part 3 |
| `wps:` | WordprocessingML Shapes | Word 2010 extension |
| `wpg:` | WordprocessingML Group | Word 2010 extension |

Namespace prefixes in actual pptx/xlsx XML files are aliases defined per-file in `xmlns` declarations. They conventionally follow this mapping, but if you see an unusual prefix, check the `xmlns` declaration in the file to confirm the actual namespace URI, then use the conventional prefix for lookup.

### How querying works

The lookup script uses a two-stage fallback — each stage only runs if the previous one returned no results:

1. **Exact name match.** Matches the element name exactly (case-sensitive, ML type prefix is optional). This is the fast path for element lookups like `a:solidFill` or `rPr`.
2. **Full-text keyword search.** Tokenizes the query into individual words (implicit AND) and searches across element names, titles, and spec text. Matches on names and titles rank higher.

Prefer prefixed element names (stage 1) for precise results. Use natural-language phrases when you don't know the element name.

If your keywords do not yield results, try reducing the number of keywords.

Exit codes: 0 = results found, 1 = no results.

## Included source documents

The ECMA-376 5th Edition (2016) specification PDFs and transitional XSD schema files are included in this skill under `pdfs/` and `schemas/` respectively. If the lookup index does not contain enough detail for your task, you may read the PDF or XSD files directly for additional context.

## Interpreting results

Each result block has this structure:

```
========================================================================
a:rPr - Text Run Properties (DrawingML)
========================================================================

This element contains all run level text properties for the text runs within a containing paragraph.
... <spec text: prose description, attributes table, XML examples, cross-references>

May appear within:
spPr, grpSpPr, ln, ...

Children:
  sequence:
    ...

Namespace: http://schemas.openxmlformats.org/drawingml/2006/main
Source: ECMA-376 Part 1, 21.1.2.3.9
```

When multiple results are returned, entries are separated by blank lines.

Key fields:
- The spec text contains the attributes table (columns may appear flattened due to PDF extraction), usage examples, and cross-references to related sections (e.g. `20.1.10.40`).
- When a body is long, a `[Match found in body]` snippet is shown instead of the full text. If you need more detail, re-run with `--limit 1` targeting the specific element.
- **May appear within:** lists the possible parent element names, derived from the XSD schema.
- **Children:** shows the content model (sequence/choice structure with cardinalities) from the XSD schema. Cardinality is shown as `[min..max]`; absent means required exactly once. `*` means `unbounded`. `...` means the nesting was truncated at depth 8.

## Caveats

**Multiple entries for the same element name.** Elements like `rPr` exist in both WordprocessingML and DrawingML with different attributes. Always check the ML type to make sure you are looking at the right spec entry.
