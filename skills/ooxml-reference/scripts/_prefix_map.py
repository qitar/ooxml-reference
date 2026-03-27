# Maps namespace prefix → (ml_type, full_namespace_uri)
PREFIX_MAP = {
    "w":   ("WordprocessingML",          "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
    "x":   ("SpreadsheetML",             "http://schemas.openxmlformats.org/spreadsheetml/2006/main"),
    "p":   ("PresentationML",            "http://schemas.openxmlformats.org/presentationml/2006/main"),
    "a":   ("DrawingML",                 "http://schemas.openxmlformats.org/drawingml/2006/main"),
    "c":   ("DrawingML Charts",          "http://schemas.openxmlformats.org/drawingml/2006/chart"),
    "dgm": ("DrawingML Diagrams",        "http://schemas.openxmlformats.org/drawingml/2006/diagram"),
    "wp":  ("DrawingML WP Drawing",      "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
    "xdr": ("DrawingML SS Drawing",      "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"),
    "cdr": ("DrawingML Chart Drawing",   "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"),
    "pic": ("DrawingML Picture",         "http://schemas.openxmlformats.org/drawingml/2006/picture"),
    "lc":  ("DrawingML Locked Canvas",   "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"),
    "r":   ("Relationships",             "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
    "mc":  ("MarkupCompatibility",       "http://schemas.openxmlformats.org/markup-compatibility/2006"),
    "m":   ("Math",                      "http://schemas.openxmlformats.org/officeDocument/2006/math"),
    "wps": ("WordprocessingML Shapes",   "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"),
    "wpg": ("WordprocessingML Group",    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"),
}


# Derived maps for quick lookups by ml_type or namespace URI.
# ml_type → prefix (first match wins)
ML_TO_PREFIX: dict[str, str] = {}
for _pfx, (_ml, _uri) in PREFIX_MAP.items():
    if _ml not in ML_TO_PREFIX:
        ML_TO_PREFIX[_ml] = _pfx

# namespace URI → prefix
URI_TO_PREFIX: dict[str, str] = {}
for _pfx, (_ml, _uri) in PREFIX_MAP.items():
    URI_TO_PREFIX.setdefault(_uri, _pfx)

# namespace URI → ml_type
URI_TO_ML: dict[str, str] = {}
for _pfx, (_ml, _uri) in PREFIX_MAP.items():
    URI_TO_ML.setdefault(_uri, _ml)


def prefix_to_ml(prefix: str):
    """Return (ml_type, namespace_uri) for a given prefix, or (None, None) if unknown."""
    entry = PREFIX_MAP.get(prefix)
    if entry:
        return entry
    return (None, None)


# Maps top-level chapter number → (ml_type, prefix) for Part 1.
# Parts 2-4 are handled directly in section_to_ml.
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
    21: ("DrawingML",        "a"),
    22: ("Shared",           None),  # mixed namespaces; subsections override below
    23: ("Overview",         None),
}

# (chapter, subsection) → (ml_type, prefix)
# Overrides CHAPTER_MAP for chapters that contain multiple sub-namespaces.
SUBSECTION_MAP = {
    # Chapter 20: DrawingML sub-namespaces
    (20, 2): ("DrawingML Chart Drawing",  "cdr"),
    (20, 3): ("DrawingML Locked Canvas",  "lc"),
    (20, 4): ("DrawingML WP Drawing",     "wp"),
    (20, 5): ("DrawingML SS Drawing",     "xdr"),
    # Chapter 21: DrawingML components
    (21, 2): ("DrawingML Charts",         "c"),
    (21, 3): ("DrawingML Charts",         "c"),
    (21, 4): ("DrawingML Diagrams",       "dgm"),
    # Chapter 22: Shared sub-namespaces
    (22, 1): ("Math",                     "m"),
    (22, 8): ("Relationships",            "r"),
}


def section_to_ml(section: str, source_part: int = 1) -> tuple[str, str | None]:
    """Return (ml_type, prefix) for a dotted section number."""
    if source_part == 2:
        return ("OpenPackagingConventions", None)
    if source_part == 3:
        return ("MarkupCompatibility", "mc")
    if source_part == 4:
        return ("Transitional", None)

    parts = section.split(".")
    chapter = int(parts[0])

    if len(parts) >= 2:
        sub = int(parts[1])
        override = SUBSECTION_MAP.get((chapter, sub))
        if override:
            return override

    return CHAPTER_MAP.get(chapter, ("Unknown", None))
