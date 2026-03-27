"""
ECMA-376 OOXML reference lookup tool.

Queries the FTS5 SQLite index built by _build_index.py. Two-stage fallback:
  1. Exact local_name match (with optional ml_type filter from namespace prefix)
  2. Tokenized FTS with snippet extraction (implicit AND; bm25 weights prioritize local_name/title)
"""

import argparse
import re
import sqlite3
import sys
from pathlib import Path

from _prefix_map import PREFIX_MAP

DB_PATH = Path(__file__).parent / "index.db"

PART_LABELS = {
    1: "ECMA-376 Part 1",
    2: "ECMA-376 Part 2",
    3: "ECMA-376 Part 3",
    4: "ECMA-376 Part 4",
}

ENTRY_SEP = "\n\n" + "-" * 72 + "\n\n"


def open_db() -> sqlite3.Connection:
    if not DB_PATH.exists():
        print(
            f"Error: index database not found at {DB_PATH}\n"
            "Run build.py first to build the index."
        )
        sys.exit(1)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def escape_fts5(term: str) -> str:
    """
    Wrap term in double-quotes so FTS5 treats it as a literal rather than
    interpreting operators like AND/OR/NOT or special chars as syntax.
    Internal double-quotes are doubled per the FTS5 spec.
    """
    return '"' + term.replace('"', '""') + '"'


def tokenize_fts5(query: str) -> str:
    """
    Split query into words and quote each one individually.
    FTS5 implicit-ANDs space-separated terms, so this behaves like
    a typical search engine: all terms must appear, in any order.
    """
    tokens = query.split()
    if not tokens:
        return '""'
    return " ".join(escape_fts5(t) for t in tokens)


def query_schema(
    conn: sqlite3.Connection, local_name: str, ml_type: str | None
) -> tuple[str | None, str | None]:
    """
    Return (parents_str, children_str) from the schema tables, or None for each
    if no data is found. Silently returns (None, None) if the tables don't exist
    (i.e. _build_schema.py has not been run yet).
    """
    try:
        if ml_type:
            row_p = conn.execute(
                "SELECT parents FROM schema_parents WHERE local_name=? AND ml_type=?",
                (local_name, ml_type),
            ).fetchone()
            row_c = conn.execute(
                "SELECT content_model FROM schema_children WHERE local_name=? AND ml_type=?",
                (local_name, ml_type),
            ).fetchone()
        else:
            row_p = conn.execute(
                "SELECT parents FROM schema_parents WHERE local_name=?",
                (local_name,),
            ).fetchone()
            row_c = conn.execute(
                "SELECT content_model FROM schema_children WHERE local_name=?",
                (local_name,),
            ).fetchone()
    except sqlite3.OperationalError:
        # Tables don't exist — _build_schema.py hasn't been run
        return None, None
    return (row_p[0] if row_p else None, row_c[0] if row_c else None)


def format_result(
    row: sqlite3.Row,
    snippet: str | None = None,
    parents: str | None = None,
    children: str | None = None,
    summary: bool = False,
) -> str:
    """
    Produce the human-readable block for a single chunk row.
    snippet is only provided for stage-2 body matches where the body is long.
    parents and children come from the schema tables built by _build_schema.py.
    """
    section = row["section"] or ""
    local_name = row["local_name"] or ""
    title = row["title"] or ""
    ml_type = row["ml_type"] or ""
    source_part = row["source_part"]
    body = row["body"] or ""

    prefix = ""
    namespace_uri = ""
    for pfx, (ml, uri) in PREFIX_MAP.items():
        if ml == ml_type:
            # Use the first matching prefix
            if not prefix:
                prefix = pfx
                namespace_uri = uri

    display_name = f"{prefix}:{local_name}" if prefix else local_name
    part_label = PART_LABELS.get(source_part, f"ECMA-376 Part {source_part}")

    lines = [
        f"=== {display_name} - {title} ===",
    ]
    if namespace_uri:
        lines.append(f"Namespace: {namespace_uri}")

    source_detail = ""
    if section:
        source_detail += f",  {section}"
    if ml_type:
        source_detail += f" ({ml_type})"
    lines.append(f"Source: {part_label}{source_detail}")
    lines.append("")

    # The body often starts with a line repeating the section number and title;
    # strip it since that info is already in the header.
    body_text = body
    if section and body_text.lstrip().startswith(section):
        body_text = body_text.split("\n", 1)[-1] if "\n" in body_text else ""

    if summary:
        first_para = body_text.split("\n\n")[0].strip()
        if first_para:
            lines.append(first_para)
    elif snippet is not None:
        lines.append("[Match found in body]")
        lines.append(snippet)
    else:
        lines.append(body_text)

    if not summary:
        if parents:
            lines.append("")
            lines.append("Parents:")
            lines.append(f"{parents}")
        if children:
            lines.append("")
            lines.append("Children:")
            for line in children.splitlines():
                lines.append("  " + line)

    return "\n".join(lines)


def stage1_exact(
    conn: sqlite3.Connection,
    local_name: str,
    ml_type: str | None,
    limit: int,
    part: int | None,
) -> list[sqlite3.Row]:
    conditions = ["local_name = ?"]
    params: list = [local_name]

    if ml_type:
        conditions.append("ml_type = ?")
        params.append(ml_type)
    if part:
        conditions.append("source_part = ?")
        params.append(part)

    order = "ORDER BY ml_type, section" if not ml_type else "ORDER BY section"
    sql = f"SELECT * FROM chunks WHERE {' AND '.join(conditions)} {order} LIMIT ?"
    params.append(limit)

    return conn.execute(sql, params).fetchall()


def stage2_fts_body(
    conn: sqlite3.Connection,
    terms: str,
    ml_type: str | None,
    limit: int,
    part: int | None,
) -> list[tuple[sqlite3.Row, str]]:
    """
    Full-text search across all columns with snippet extraction from body (column 2).
    bm25 weights (10, 5, 1) ensure matches on local_name or title rank above body-only hits.
    """
    fts_query = terms
    ml_clause = "AND c.ml_type = ?" if ml_type else ""
    part_clause = "AND c.source_part = ?" if part else ""

    sql = f"""
        SELECT c.*, snippet(chunks_fts, 2, '>>>', '<<<', '...', 32) AS body_snippet
        FROM chunks_fts
        JOIN chunks c ON chunks_fts.rowid = c.id
        WHERE chunks_fts MATCH ?
        {ml_clause}
        {part_clause}
        ORDER BY bm25(chunks_fts, 10.0, 5.0, 1.0)
        LIMIT ?
    """
    params: list = [fts_query]
    if ml_type:
        params.append(ml_type)
    if part:
        params.append(part)
    params.append(limit)

    try:
        rows = conn.execute(sql, params).fetchall()
    except sqlite3.OperationalError:
        return []

    return [(row, row["body_snippet"]) for row in rows]


def print_no_results(query: str, local_name: str) -> None:
    print(f'No results found for "{query}".')
    print("Suggestions:")
    if ":" in query:
        print(f'- Try without the namespace prefix: "{local_name}"')
    print(f'- Try a descriptive phrase: "{local_name} properties"')
    print("- Check that the index has been built by verifying index.db exists")


def lookup(query: str, limit: int, part: int | None, summary: bool = False) -> bool:
    """Run the two-stage lookup and print results. Returns True if results were found."""
    conn = open_db()

    # Parse optional namespace prefix
    ml_type: str | None = None
    local_name = query
    if ":" in query:
        prefix, _, rest = query.partition(":")
        # Only treat as a namespace prefix if it looks like one (no spaces, maps to something)
        if prefix and not re.search(r"\s", prefix):
            local_name = rest
            ml_type, _ = PREFIX_MAP.get(prefix, (None, None))

    def fmt(row: sqlite3.Row, snippet: str | None = None) -> str:
        if summary:
            return format_result(row, snippet, summary=True)
        parents, children = query_schema(conn, row["local_name"] or "", row["ml_type"])
        return format_result(row, snippet, parents, children)

    with conn:
        # Stage 1 — exact local_name match
        rows = stage1_exact(conn, local_name, ml_type, limit, part)

        if rows:
            print(ENTRY_SEP.join(fmt(r) for r in rows))
            return True

        # Stage 2 — full-text search with snippets (bm25 weights prioritize local_name/title)
        results = stage2_fts_body(conn, tokenize_fts5(query), ml_type, limit, part)

        if results:
            blocks = [fmt(row, snippet) for row, snippet in results]
            print(ENTRY_SEP.join(blocks))
            return True

    print_no_results(query, local_name)
    return False


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Look up ECMA-376 OOXML spec entries from the FTS index."
    )
    parser.add_argument(
        "query",
        help="Search term, e.g. 'w:rPr', 'solidFill', 'bold text'",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=5,
        metavar="N",
        help="Maximum number of results (default: 5)",
    )
    parser.add_argument(
        "--part",
        type=int,
        choices=[1, 2, 3, 4],
        default=None,
        help="Restrict results to a specific ECMA-376 source part",
    )
    parser.add_argument(
        "--summary",
        action="store_true",
        help="Show only the section title, namespace, and first paragraph",
    )
    args = parser.parse_args()

    found = lookup(args.query, args.limit, args.part, summary=args.summary)
    if not found:
        sys.exit(1)


if __name__ == "__main__":
    main()
