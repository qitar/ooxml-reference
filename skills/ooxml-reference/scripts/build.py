#!/usr/bin/env python3
"""Rebuild the full OOXML spec index (FTS + schema tables)."""

from _build_index import main as build_index
from _build_schema import main as build_schema

if __name__ == "__main__":
    build_index()
    build_schema()
