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
from collections import defaultdict
from pathlib import Path

from _prefix_map import section_to_ml

PDFS = {
    1: "ECMA-376 OOXML (1) Fundamentals And Markup Language Reference.pdf",
    2: "ECMA-376 OOXML (2) Open Packaging Conventions.pdf",
    3: "ECMA-376 OOXML (3) Markup Compatibility and Extensibility.pdf",
    4: "ECMA-376 OOXML (4) Transitional Migration Features.pdf",
}

DB_PATH = Path(__file__).parent / "index.db"
PDF_DIR = Path(__file__).parent.parent / "pdfs"
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
        raise RuntimeError("pdftotext not found. Install it with: brew install poppler")


def extract_pdf(pdf_path: Path, txt_path: Path):
    result = subprocess.run(
        ["pdftotext", "-layout", str(pdf_path), str(txt_path)],
        capture_output=True,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"pdftotext failed on {pdf_path} (exit {result.returncode}):\n"
            f"{result.stderr.decode(errors='replace')}"
        )


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

        ml_type, prefix = section_to_ml(current_section, source_part)
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
    try:
        check_pdftotext()

        TMP_DIR.mkdir(parents=True, exist_ok=True)

        conn = init_db(DB_PATH)

        counts = defaultdict(lambda: defaultdict(int))
        all_chunks = []

        for i, (part, path) in enumerate(PDFS.items()):
            pdf_path = PDF_DIR / path
            txt_path = TMP_DIR / f"part{part}.txt"

            print(f"[Part {part}/{len(PDFS)}] ", end="")
            if not pdf_path.exists():
                raise FileNotFoundError(f"Part {part} PDF not found at {pdf_path}")

            print("Extracting text... ", end="", flush=True)
            extract_pdf(pdf_path, txt_path)
            print("done")

            print("           Parsing text...    ", end="", flush=True)
            part_chunks = list(parse_chunks(txt_path, part))
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
    finally:
        shutil.rmtree(TMP_DIR, ignore_errors=True)
