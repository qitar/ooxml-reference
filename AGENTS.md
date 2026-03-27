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
| `skills/ooxml-reference/scripts/_prefix_map.py` | Namespace prefix/URI/ML-type mappings and section→ML-type resolution |
| `skills/ooxml-reference/scripts/index.db` | SQLite database (FTS5 index + schema tables) |
| `DESIGN.md` | Architecture and implementation details |

## Critical invariant

The `ml_type` values must be consistent across two places:

1. **`_prefix_map.py`** — `PREFIX_MAP` maps prefixes to (ml_type, namespace_uri) tuples;
   `CHAPTER_MAP` and `SUBSECTION_MAP` assign ml_type to PDF sections via `section_to_ml`
2. **`_build_schema.py`** — derives ml_type from `PREFIX_MAP` via namespace URIs (`URI_TO_ML`)

Both data paths derive from `PREFIX_MAP`, so they stay in sync as long as the ml_type
strings in `SUBSECTION_MAP` match those in `PREFIX_MAP`.

If these diverge, prefixed lookups (e.g. `c:barChart`) silently return no results because
the stage-1 exact match filters on ml_type.

## Rebuilding the index

Run from `skills/ooxml-reference/scripts/`:

```bash
uv run --with pymupdf python build.py
```

The index must be rebuilt after any change to `_build_index.py`, `_prefix_map.py`, or `_build_schema.py`.
