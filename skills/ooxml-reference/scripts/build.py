#!/usr/bin/env python3
"""Rebuild the full OOXML spec index (FTS + schema tables)."""

from pathlib import Path

from _build_index import main as build_index
from _build_schema import main as build_schema

DB_PATH = Path(__file__).parent / "index.db"

if __name__ == "__main__":
    if DB_PATH.exists():
        DB_PATH.unlink()
        print(f"Dropped existing index at {DB_PATH}")
        print("")

    build_index()

    print("\n" + "-"*72 + "\n")

    build_schema()
