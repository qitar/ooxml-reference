# Maps namespace prefix → (ml_type, full_namespace_uri)
# Chapter ranges are for Part 1 section numbers (top-level chapter = first digit(s) before first dot)
PREFIX_MAP = {
    "w":   ("WordprocessingML",        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
    "x":   ("SpreadsheetML",           "http://schemas.openxmlformats.org/spreadsheetml/2006/main"),
    "p":   ("PresentationML",          "http://schemas.openxmlformats.org/presentationml/2006/main"),
    "a":   ("DrawingML",               "http://schemas.openxmlformats.org/drawingml/2006/main"),
    "c":   ("DrawingML Charts",        "http://schemas.openxmlformats.org/drawingml/2006/chart"),
    "dgm": ("DrawingML Diagrams",      "http://schemas.openxmlformats.org/drawingml/2006/diagram"),
    "r":   ("Relationships",           "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
    "mc":  ("MarkupCompatibility",     "http://schemas.openxmlformats.org/markup-compatibility/2006"),
    "m":   ("Math",                    "http://schemas.openxmlformats.org/officeDocument/2006/math"),
    "wps": ("WordprocessingML Shapes", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"),
    "wpg": ("WordprocessingML Group",  "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"),
}

# Maps top-level chapter number (int) → (ml_type, prefix)
# For Part 1 only; Parts 2-4 get their own source_part tag.
# Chapter 21 has subsections that use c: and dgm: — refined in _build_index.py.
CHAPTER_MAP = {
    # Introductory/overview chapters (before the reference material)
    2:  ("Overview", None),
    3:  ("Overview", None),
    4:  ("Overview", None),
    5:  ("Overview", None),
    6:  ("Overview", None),
    7:  ("Overview", None),
    8:  ("Overview", None),
    9:  ("Overview", None),
    10: ("Overview", None),
    11: ("WordprocessingML", "w"),   # package/part structure overview
    12: ("SpreadsheetML",    "x"),
    13: ("PresentationML",   "p"),
    14: ("DrawingML",        "a"),
    15: ("Shared",           None),
    16: ("Overview",         None),
    # Reference material chapters
    17: ("WordprocessingML", "w"),
    18: ("SpreadsheetML",    "x"),
    19: ("PresentationML",   "p"),
    20: ("DrawingML",        "a"),
    21: ("DrawingML",        "a"),   # components; subsections refined to c: and dgm: in _build_index.py
    22: ("Shared",           None),  # mixed namespaces; no single prefix
    23: ("Overview",         None),
}


def prefix_to_ml(prefix: str):
    """Return (ml_type, namespace_uri) for a given prefix, or (None, None) if unknown."""
    entry = PREFIX_MAP.get(prefix)
    if entry:
        return entry
    return (None, None)


def chapter_to_ml(chapter: int, source_part: int = 1):
    """Return (ml_type, prefix) for a top-level chapter number."""
    if source_part == 2:
        return ("OpenPackagingConventions", None)
    if source_part == 3:
        return ("MarkupCompatibility", "mc")
    if source_part == 4:
        return ("Transitional", None)
    return CHAPTER_MAP.get(chapter, ("Unknown", None))
