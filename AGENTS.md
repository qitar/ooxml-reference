# Agent Guide

## What this repo is

A Claude Code skill that provides authoritative ECMA-376 OOXML specification lookup.
The skill lets an AI agent query element definitions, attributes, parent/child relationships,
and content models from the OOXML standard (used by `.pptx`, `.xlsx`, `.docx` files).

## Key files

| File | Purpose |
|------|---------|
| `skills/ooxml-reference/SKILL.md` | Skill definition read by the agent at invocation time |
| `skills/ooxml-reference/scripts/lookup.py` | Entry point script invoked by the skill |
| `skills/ooxml-reference/scripts/query.py` | Query logic: three-stage FTS fallback |
| `skills/ooxml-reference/scripts/build_index.py` | Builds the FTS index from ECMA-376 PDFs |
| `skills/ooxml-reference/scripts/build_schema.py` | Builds parent/child tables from XSD schemas |
| `skills/ooxml-reference/scripts/prefix_map.py` | Canonical namespace prefix → ML type mapping |
| `skills/ooxml-reference/scripts/index.db` | SQLite database (FTS5 index + schema tables) |
| `DESIGN.md` | Architecture and implementation details |

## Critical invariant

The `ml_type` values must be consistent across three places:

1. **`prefix_map.py`** — `PREFIX_MAP` maps prefixes to ml_type strings
2. **`build_index.py`** — assigns ml_type when chunking PDFs (via `chapter_to_ml` + `refine_ml_and_prefix`)
3. **`build_schema.py`** — assigns ml_type when parsing XSD files (via `XSD_ML` dict)

If these diverge, prefixed lookups (e.g. `c:barChart`) silently return no results because
the stage-1 exact match filters on ml_type.

## Rebuilding the index

Both scripts must be run in order from `skills/ooxml-reference/scripts/`:

```bash
python build_index.py    # requires pdftotext (brew install poppler)
python build_schema.py
```

The index must be rebuilt after any change to `build_index.py`, `prefix_map.py`, or `build_schema.py`.
