"""
Build the OOXML spec index from ECMA-376 PDFs.

Writes to tables in index.db: chunks and chunks_fts.

This script is tailor-made for the specific PDFs.
Updates to the spec will likely require an entirely new parsing logic.
"""

import re
import sqlite3
from collections import defaultdict
from pathlib import Path

import fitz  # type: ignore

from _prefix_map import section_to_ml

PDFS = {
    1: "ECMA-376 OOXML (1) Fundamentals And Markup Language Reference.pdf",
    2: "ECMA-376 OOXML (2) Open Packaging Conventions.pdf",
    3: "ECMA-376 OOXML (3) Markup Compatibility and Extensibility.pdf",
    4: "ECMA-376 OOXML (4) Transitional Migration Features.pdf",
}

DB_PATH = Path(__file__).parent / "index.db"
PDF_DIR = Path(__file__).parent.parent / "pdfs"

# All section headings use Cambria (regular or Bold) at >= 12pt, while body text
# uses Calibri 11pt and code uses Consolas 10-11pt.
HEADING_MIN_SIZE = 12.0

# Page headers (running titles) and footers (page numbers) sit at fixed y-positions.
# These thresholds (in points) exclude them from content extraction.
MARGIN_TOP = 70
MARGIN_BOTTOM = 730

# Applied only to text already confirmed as heading-font. Extracts section number,
# element name, and parenthesized title. No anti-false-positive guards needed here
# since font metadata has already confirmed it's a real heading.
SECTION_RE = re.compile(
    r"^\s*(\d+(?:\.\d+)+)\s+(\S.*?)(?:\s*\((.+?)\))?\s*$"
)

# PDF tables that span page breaks repeat the header row on each new page.
# We keep only the first occurrence per chunk.
TABLE_HEADER_RE = re.compile(r"^\s*Attributes\s{2,}Description\s*$")

# Fixed character width for the attribute-name column in formatted output.
ATTR_COL_WIDTH = 24


def is_heading_span(span: dict) -> bool:
    """True when a span uses Cambria at heading size (>= 12pt)."""
    return "Cambria" in span["font"] and span["size"] >= HEADING_MIN_SIZE


def is_heading_block(block: dict) -> bool:
    """True when any line in a text block contains a heading-styled span.

    Heading blocks often have two lines (section number + title) where the
    section-number line uses a different style (Cambria-Bold 11pt black).
    Checking at block level catches these mixed-style headings.
    """
    return any(
        is_heading_span(line["spans"][0])
        for line in block["lines"]
        if line["spans"]
    )


def _spans_to_text(spans: list[dict]) -> str:
    """Join spans into text, inserting spaces to preserve column gaps.

    pymupdf spans carry bbox coordinates. When two consecutive spans are
    separated by a gap wider than a few normal characters, the original PDF
    had a column or table boundary there. We approximate the gap with spaces
    so that attribute tables remain readable in plain text.
    """
    if not spans:
        return ""

    parts = [spans[0]["text"]]
    for prev, cur in zip(spans, spans[1:]):
        gap = cur["bbox"][0] - prev["bbox"][2]
        char_w = cur["size"] * 0.5
        if char_w > 0 and gap > char_w * 3:
            n_spaces = max(2, round(gap / char_w))
            parts.append(" " * n_spaces)
        parts.append(cur["text"])

    return "".join(parts)


# Sentinel value for table_col_x: header was found but the column boundary
# hasn't been determined yet (waiting for the first body row).
_SENTINEL = -1.0


def _find_col_boundary(spans: list[dict]) -> float | None:
    """Find the column boundary from the largest inter-span gap in a row.

    Returns the midpoint of the widest gap, or None if no gap exceeds 20pt
    (i.e. the row doesn't look like a two-column table row).
    """
    max_gap = 0.0
    boundary = None
    for prev, cur in zip(spans, spans[1:]):
        gap = cur["bbox"][0] - prev["bbox"][2]
        if gap > max_gap:
            max_gap = gap
            boundary = (prev["bbox"][2] + cur["bbox"][0]) / 2
    if boundary is not None and max_gap > 20:
        return boundary
    return None


def _format_table_row(spans: list[dict], col_x: float) -> str:
    """Format a row of spans as a two-column table line."""
    left = [s for s in spans if s["bbox"][0] < col_x]
    right = [s for s in spans if s["bbox"][0] >= col_x]
    left_text = _spans_to_text(left).strip() if left else ""
    right_text = _spans_to_text(right).strip() if right else ""
    return f"{left_text:<{ATTR_COL_WIDTH}}{right_text}"


def _group_rows(blocks: list[dict]) -> list[list[dict]]:
    """Collect all spans from blocks and group into rows by vertical position.

    PDF attribute tables often place the name cell and description cell in
    separate blocks that overlap vertically (same y-position, different x).
    pymupdf can also split side-by-side cells into separate "lines" within
    one block. We collect all lines from the given blocks, group them by
    vertical position, then return each row as a list of pymupdf line dicts.
    """
    all_lines = []
    for block in blocks:
        for line in block["lines"]:
            if line["spans"]:
                all_lines.append(line)

    if not all_lines:
        return []

    def y_mid(line: dict) -> float:
        return (line["bbox"][1] + line["bbox"][3]) / 2

    # Sort by vertical midpoint so lines from different blocks interleave correctly
    all_lines.sort(key=y_mid)

    # Group lines whose vertical midpoints are within half a line-height
    rows: list[list[dict]] = []
    for line in all_lines:
        if rows:
            last_row = rows[-1]
            ref_mid = y_mid(last_row[0])
            height = last_row[0]["bbox"][3] - last_row[0]["bbox"][1]
            if abs(y_mid(line) - ref_mid) < max(height * 0.5, 3.0):
                last_row.append(line)
                continue
        rows.append([line])

    return rows


def merge_block_lines(
    blocks: list[dict], table_col_x: float | None = None
) -> tuple[list[str], float | None]:
    """Merge lines across multiple pymupdf blocks, preserving column layout.

    When an attribute table is detected (via its "Attributes / Description"
    header row), spans are split into left and right columns using the
    header's x-coordinate as the boundary. This produces consistently
    aligned output without gap-based guessing.

    Returns (lines, table_col_x) so callers can carry table state across
    page breaks within a single chunk.
    """
    rows = _group_rows(blocks)
    if not rows:
        return [], table_col_x

    result: list[str] = []
    for row in rows:
        all_spans: list[dict] = []
        for line in row:
            all_spans.extend(line["spans"])
        all_spans.sort(key=lambda s: s["bbox"][0])

        if table_col_x is None:
            text = _spans_to_text(all_spans)
            if TABLE_HEADER_RE.match(text):
                # Mark that we've seen the header; the actual column boundary
                # will be determined from the first body row whose layout is
                # more reliable than the header text placement.
                table_col_x = _SENTINEL
                result.append("")
                result.append(f"{'Attributes':<{ATTR_COL_WIDTH}}Description")
            else:
                result.append(text)

        elif table_col_x == _SENTINEL:
            # First body row after the header: find the largest inter-span
            # gap and use its midpoint as the column boundary.
            table_col_x = _find_col_boundary(all_spans)
            if table_col_x is not None:
                result.append(_format_table_row(all_spans, table_col_x))
            else:
                # No clear two-column structure; not actually a table
                result.append(_spans_to_text(all_spans))

        else:
            # Inside an attribute table: use the column boundary to split spans
            text = _spans_to_text(all_spans)
            if TABLE_HEADER_RE.match(text):
                continue  # duplicate header at page break

            formatted = _format_table_row(all_spans, table_col_x)
            left_only = all(s["bbox"][0] < table_col_x for s in all_spans)

            if left_only and len(text.strip()) > ATTR_COL_WIDTH:
                # Full-width text after the table (e.g. "[Note: ..." paragraph)
                table_col_x = None
                result.append(text.strip())
            else:
                result.append(formatted)

    return result, table_col_x


def parse_chunks(pdf_path: Path, source_part: int):
    """
    Yield dicts with keys: section, local_name, title, ml_type, prefixes, source_part, body.

    Uses pymupdf font metadata to detect section headings instead of regex on
    plain text. Chunks with too little body content or unmapped ml_type are dropped.
    """
    doc = fitz.open(pdf_path)

    current_section = None
    current_local_name = None
    current_title = None
    body_lines: list[str] = []
    table_col_x: float | None = None

    def flush_chunk():
        if current_section is None:
            return None

        body = "\n".join(body_lines).strip()

        if len(body) < 50:
            return None

        ml_type, prefix = section_to_ml(current_section, source_part)

        if ml_type == "Unknown":
            return None

        prefixes = f"{prefix}:" if prefix else None

        return {
            "section": current_section,
            "local_name": current_local_name,
            "title": current_title,
            "ml_type": ml_type,
            "prefixes": prefixes,
            "source_part": source_part,
            "body": body,
        }

    for page in doc:
        blocks = page.get_text("dict")["blocks"]

        # Batch consecutive non-heading blocks so cross-block table cells
        # (attribute name in one block, description in another at the same y)
        # get their lines merged into proper rows.
        pending_body_blocks: list[dict] = []

        def flush_body_blocks():
            nonlocal table_col_x
            if pending_body_blocks:
                lines, table_col_x = merge_block_lines(
                    pending_body_blocks, table_col_x
                )
                body_lines.extend(lines)
                pending_body_blocks.clear()

        for block in blocks:
            if block["type"] != 0:
                continue

            y0 = block["bbox"][1]
            if y0 < MARGIN_TOP or y0 > MARGIN_BOTTOM:
                continue

            if is_heading_block(block):
                flush_body_blocks()

                heading_parts = [
                    _spans_to_text(line["spans"]).strip()
                    for line in block["lines"]
                    if line["spans"]
                ]
                heading_text = " ".join(heading_parts)
                m = SECTION_RE.match(heading_text)
                if m:
                    chunk = flush_chunk()
                    if chunk:
                        yield chunk

                    table_col_x = None
                    current_section = m.group(1)
                    if m.group(3):
                        current_local_name = m.group(2).strip()
                        current_title = m.group(3).strip()
                    else:
                        current_local_name = None
                        current_title = m.group(2).strip()

                    # Include the heading text in body for self-contained context
                    body_lines = [heading_text]
                else:
                    # Heading-styled block without a section number (e.g. "Foreword")
                    for part in heading_parts:
                        body_lines.append(part)
            else:
                pending_body_blocks.append(block)

        flush_body_blocks()

    # Flush the final chunk
    chunk = flush_chunk()
    if chunk:
        yield chunk

    doc.close()


def init_db(path: Path):
    conn = sqlite3.connect(path)

    conn.executescript("""
        DROP TABLE IF EXISTS chunks;
        CREATE TABLE chunks (
            id          INTEGER PRIMARY KEY,
            section     TEXT,
            local_name  TEXT,
            title       TEXT,
            ml_type     TEXT,
            prefixes    TEXT,
            source_part INTEGER,
            body        TEXT
        );

        DROP TABLE IF EXISTS chunks_fts;
        CREATE VIRTUAL TABLE chunks_fts USING fts5(
            local_name,
            title,
            body,
            content='chunks',
            content_rowid='id'
        );
    """)
    conn.commit()

    return conn


def insert_chunks(conn: sqlite3.Connection, chunks):
    conn.executemany(
        """
        INSERT INTO chunks (section, local_name, title, ml_type, prefixes, source_part, body)
        VALUES (:section, :local_name, :title, :ml_type, :prefixes, :source_part, :body)
        """,
        chunks,
    )
    conn.commit()


def populate_fts(conn: sqlite3.Connection):
    conn.execute("""
        INSERT INTO chunks_fts(rowid, local_name, title, body)
        SELECT id, local_name, title, body FROM chunks
    """)
    conn.commit()


def main():
    conn = init_db(DB_PATH)

    counts = defaultdict(lambda: defaultdict(int))
    all_chunks = []

    for part, path in PDFS.items():
        pdf_path = PDF_DIR / path

        print(f"[Part {part}/{len(PDFS)}] ", end="")
        if not pdf_path.exists():
            raise FileNotFoundError(f"Part {part} PDF not found at {pdf_path}")

        print("Parsing... ", end="", flush=True)
        part_chunks = list(parse_chunks(pdf_path, part))
        print(f"done ({len(part_chunks)} chunks)")

        for chunk in part_chunks:
            counts[part][chunk["ml_type"]] += 1

        all_chunks.extend(part_chunks)

    print(f"\nInserting {len(all_chunks)} chunks into database... ", end="", flush=True)
    insert_chunks(conn, all_chunks)
    populate_fts(conn)
    print("done")

    conn.close()

    print("\nPDF indexing complete:")
    total = 0
    for part in sorted(counts):
        for ml_type, count in sorted(counts[part].items()):
            label = f"Part {part} - {ml_type}"
            print(f"  {label:<40} {count} chunks")
            total += count
    print(f"  {'Total':<40} {total} chunks")
