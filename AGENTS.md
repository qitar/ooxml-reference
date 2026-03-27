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
| `skills/ooxml-reference/scripts/_query.py` | Query logic: two-stage FTS fallback |
| `skills/ooxml-reference/scripts/build.py` | Entry point to rebuild the full index |
| `skills/ooxml-reference/scripts/_build_index.py` | Builds the FTS index from ECMA-376 PDFs |
| `skills/ooxml-reference/scripts/_build_schema.py` | Builds parent/child tables from XSD schemas |
| `skills/ooxml-reference/scripts/_prefix_map.py` | Canonical namespace prefix → ML type mapping |
| `skills/ooxml-reference/scripts/index.db` | SQLite database (FTS5 index + schema tables) |
| `DESIGN.md` | Architecture and implementation details |

## Critical invariant

The `ml_type` values must be consistent across two places:

1. **`_prefix_map.py`** — `PREFIX_MAP` maps prefixes to (ml_type, namespace_uri) tuples
2. **`_build_index.py`** — assigns ml_type when chunking PDFs (via `chapter_to_ml` + `refine_ml_and_prefix`)

`_build_schema.py` derives its ml_type values from `PREFIX_MAP` via namespace URIs
(`URI_TO_ML`), so it stays in sync automatically. DrawingML sub-namespace URIs
(chartDrawing, picture, etc.) are mapped explicitly in `_DML_SUB_NS`.

If these diverge, prefixed lookups (e.g. `c:barChart`) silently return no results because
the stage-1 exact match filters on ml_type.

## Rebuilding the index

Run from `skills/ooxml-reference/scripts/`:

```bash
python build.py    # requires pdftotext (brew install poppler)
```

The index must be rebuilt after any change to `_build_index.py`, `_prefix_map.py`, or `_build_schema.py`.
