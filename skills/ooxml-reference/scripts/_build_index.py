"""
Build the OOXML spec index from ECMA-376 PDFs.

Writes to tables in index.db: chunks and chunks_fts.

This script is tailor-made for the specific PDFs.
Updates to the spec will likely require an entirely new parsing logic.
"""

import re
import shutil
import sqlite3
import subprocess
import sys
import time
from collections import defaultdict
from pathlib import Path

from _prefix_map import chapter_to_ml

PDFS = {
    1: "ECMA-376 OOXML (1) Fundamentals And Markup Language Reference.pdf",
    2: "ECMA-376 OOXML (2) Open Packaging Conventions.pdf",
    3: "ECMA-376 OOXML (3) Markup Compatibility and Extensibility.pdf",
    4: "ECMA-376 OOXML (4) Transitional Migration Features.pdf",
}
PDF_PAGE_COUNTS = {1: 5039, 2: 95, 3: 43, 4: 1553}

PDF_DIR = Path(__file__).parent.parent / "pdfs"
DB_PATH = Path(__file__).parent / "index.db"
TMP_DIR = Path(__file__).parent / "tmp"

# Matches lines like: "17.3.2.27    rPr (Run Properties)"
# Group 1: section number, Group 2: local name, Group 3: title (optional)
# Requires chapter >= 1 to avoid matching body-text decimals like "0.25 inches".
HEADING_RE = re.compile(
    r"^\s*([1-9]\d*(?:\.\d+)+)\s+(\S.*?)(?:\s*\((.+?)\))?\s*$"
)

# TOC heading lines have dotted leaders followed by a page number at the end.
# These match HEADING_RE but should not start a new chunk — they're navigation
# artifacts that would create duplicate chunks alongside the real content pages.
TOC_HEADING_RE = re.compile(r"\.{3,}\s*\d+\s*$")

# TOC entries have no real body — just dotted leaders and a page number
TOC_BODY_RE = re.compile(r"[\s.·\-–—]*\d*[\s.·\-–—]*")


def check_pdftotext():
    if not shutil.which("pdftotext"):
        print("Error: pdftotext not found. Install it with: brew install poppler")
        sys.exit(1)


def extract_pdf(pdf_path: Path, txt_path: Path, part: int):
    page_count = PDF_PAGE_COUNTS.get(part, "?")
    print(f"  Extracting Part {part} ({page_count} pages)...", end=" ", flush=True)
    t0 = time.time()
    result = subprocess.run(
        ["pdftotext", "-layout", str(pdf_path), str(txt_path)],
        capture_output=True,
    )
    elapsed = int(time.time() - t0)
    if result.returncode != 0:
        print(f"FAILED (exit {result.returncode})")
        print(result.stderr.decode(errors="replace"))
        return False
    print(f"done in {elapsed}s")
    return True


def refine_ml_and_prefix(
    section: str, default_ml: str, default_prefix: str | None
) -> tuple[str, str | None]:
    """
    Some chapters cover multiple sub-namespaces. Narrow both ml_type and prefix
    based on the subsection number so they match PREFIX_MAP exactly.
    """
    # Chapter 21: DrawingML sub-namespaces
    if section.startswith("21.2") or section.startswith("21.3"):
        return "DrawingML Charts", "c:"
    if section.startswith("21.4"):
        return "DrawingML Diagrams", "dgm:"
    # Chapter 22: Shared sub-namespaces
    if section.startswith("22.1"):
        return "Math", "m:"
    if section.startswith("22.8"):
        return "Relationships", "r:"

    prefix = f"{default_prefix}:" if default_prefix else None
    return default_ml, prefix


def strip_page_margins(text: str, header_lines: int = 1, footer_lines: int = 1) -> str:
    """
    Remove running page headers and footers from pdftotext -layout output.

    pdftotext delimits pages with form-feed characters. Headers and footers
    sit at fixed positions at the top and bottom of each page, so slicing
    a fixed number of non-blank lines from each end reliably removes them
    without brittle pattern matching.
    """
    pages = text.split("\f")
    stripped = []
    for page in pages:
        page_lines = page.splitlines()
        # Collect indices of non-blank lines so we skip blank padding
        non_blank = [i for i, ln in enumerate(page_lines) if ln.strip()]
        if len(non_blank) <= header_lines + footer_lines:
            # Page has too few real lines to strip — keep as-is so we don't
            # accidentally discard a real content page with a short section.
            stripped.append(page)
            continue
        drop_top = set(non_blank[:header_lines])
        drop_bot = set(non_blank[-footer_lines:])
        kept = [ln for i, ln in enumerate(page_lines) if i not in drop_top and i not in drop_bot]
        stripped.append("\n".join(kept))
    return "\f".join(stripped)


def parse_chunks(txt_path: Path, source_part: int):
    """
    Yield dicts with keys: section, local_name, title, ml_type, prefixes, source_part, body.

    Chunks that are TOC entries or too short are silently dropped.
    """
    text = txt_path.read_text(encoding="utf-8", errors="replace")
    text = strip_page_margins(text)
    lines = text.splitlines()

    current_section = None
    current_local_name = None
    current_title = None
    body_lines = []

    def flush_chunk():
        if current_section is None:
            return None
        body = "\n".join(body_lines).strip()

        # Drop TOC entries: body is blank or only dots/dashes/page-number
        if not body or TOC_BODY_RE.fullmatch(body.strip()):
            return None

        # Drop sections with too little content to be useful
        if len(body.strip()) < 50:
            return None

        chapter = int(current_section.split(".")[0])
        ml_type, raw_prefix = chapter_to_ml(chapter, source_part)

        # Refine ml_type and prefix for chapters with sub-namespaces
        if source_part == 1:
            ml_type, prefixes = refine_ml_and_prefix(
                current_section, ml_type, raw_prefix
            )
        elif source_part == 3:
            prefixes = "mc:"
        else:
            prefixes = f"{raw_prefix}:" if raw_prefix else None

        return {
            "section": current_section,
            "local_name": current_local_name,
            "title": current_title,
            "ml_type": ml_type,
            "prefixes": prefixes,
            "source_part": source_part,
            "body": body,
        }

    for line in lines:
        m = HEADING_RE.match(line)
        if m and not TOC_HEADING_RE.search(line):
            chunk = flush_chunk()
            if chunk:
                yield chunk

            current_section = m.group(1)
            # The local name may be followed by the title in parens on the same line;
            # group(2) is everything between the section number and the optional paren group.
            current_local_name = m.group(2).strip()
            current_title = m.group(3).strip() if m.group(3) else None
            # The heading line itself is part of the body so context is self-contained
            body_lines = [line]
        elif current_section is not None:
            body_lines.append(line)

    # Flush the final chunk
    chunk = flush_chunk()
    if chunk:
        yield chunk


def init_db(conn: sqlite3.Connection):
    conn.executescript("""
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

        CREATE VIRTUAL TABLE chunks_fts USING fts5(
            local_name,
            title,
            body,
            content='chunks',
            content_rowid='id'
        );
    """)
    conn.commit()


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


def print_summary(counts: dict):
    print("\nIndexing complete:")
    total = 0
    for part in sorted(counts):
        for ml_type, count in sorted(counts[part].items()):
            label = f"Part {part} — {ml_type}:"
            print(f"  {label:<40} {count} chunks")
            total += count
    print(f"  {'Total:':<40} {total} chunks")


def main():
    check_pdftotext()
    TMP_DIR.mkdir(parents=True, exist_ok=True)

    if DB_PATH.exists():
        DB_PATH.unlink()
        print(f"Dropped existing index at {DB_PATH}")

    parts_to_index = [1, 2, 3, 4]
    total_parts = len(parts_to_index)

    conn = sqlite3.connect(DB_PATH)
    init_db(conn)

    counts = defaultdict(lambda: defaultdict(int))
    all_chunks = []

    for i, part in enumerate(parts_to_index, 1):
        pdf_path = PDF_DIR / PDFS[part]
        txt_path = TMP_DIR / f"part{part}.txt"

        print(f"[{i}/{total_parts}] ", end="")

        if not pdf_path.exists():
            print(f"Warning: Part {part} PDF not found at {pdf_path}, skipping.")
            continue

        if not extract_pdf(pdf_path, txt_path, part):
            continue

        print(f"  Parsing Part {part}...", end=" ", flush=True)
        t0 = time.time()
        part_chunks = list(parse_chunks(txt_path, part))
        elapsed = int(time.time() - t0)
        print(f"done in {elapsed}s ({len(part_chunks)} chunks)")

        for chunk in part_chunks:
            counts[part][chunk["ml_type"]] += 1
        all_chunks.extend(part_chunks)

    print(f"\nInserting {len(all_chunks)} chunks into database...", end=" ", flush=True)
    t0 = time.time()
    insert_chunks(conn, all_chunks)
    populate_fts(conn)
    conn.close()
    print(f"done in {int(time.time() - t0)}s")

    print_summary(counts)
