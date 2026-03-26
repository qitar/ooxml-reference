"""
Build the schema parent/child index from ECMA-376 XSD files.

Parses XSD files to extract element→children and element→parents relationships,
then stores them in two tables in index.db. The transitional XSD set is used by
default because its namespace URIs match PREFIX_MAP in prefix_map.py.

Usage:
    python skills/ooxml/build_schema.py
    python skills/ooxml/build_schema.py --xsd-dir strict
"""

import argparse
import sqlite3
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from prefix_map import PREFIX_MAP  # noqa: E402

DB_PATH = Path(__file__).parent / "index.db"
SOURCE_DIR = Path(__file__).parent.parent.parent / "source_docs"

# Reverse map: ml_type → namespace prefix (first match wins)
ML_TO_PREFIX: dict[str, str] = {}
for _pfx, (_ml, _uri) in PREFIX_MAP.items():
    if _ml not in ML_TO_PREFIX:
        ML_TO_PREFIX[_ml] = _pfx

XSD_NS = "http://www.w3.org/2001/XMLSchema"

# Maps XSD filename → ML type string matching the chunks table's ml_type column.
XSD_ML = {
    "wml.xsd": "WordprocessingML",
    "pml.xsd": "PresentationML",
    "sml.xsd": "SpreadsheetML",
    "dml-main.xsd": "DrawingML",
    "dml-chart.xsd": "DrawingML Charts",
    "dml-diagram.xsd": "DrawingML Diagrams",
    "dml-chartDrawing.xsd": "DrawingML",
    "dml-lockedCanvas.xsd": "DrawingML",
    "dml-picture.xsd": "DrawingML",
    "dml-spreadsheetDrawing.xsd": "DrawingML",
    "dml-wordprocessingDrawing.xsd": "DrawingML",
    "shared-math.xsd": "Math",
    "shared-relationshipReference.xsd": "Relationships",
}


def file_to_ml(filename: str) -> str:
    ml = XSD_ML.get(filename)
    if ml:
        return ml
    if filename.startswith("opc-"):
        return "OpenPackagingConventions"
    return "Shared"


def local(tag: str) -> str:
    """Strip the XSD namespace URI from a qualified ElementTree tag."""
    return tag.split("}")[1] if "}" in tag else tag


def parse_occurs(elem) -> tuple[int, str]:
    min_o = int(elem.get("minOccurs", "1"))
    max_o = elem.get("maxOccurs", "1")
    return min_o, max_o


def parse_node(elem) -> dict | None:
    """
    Recursively parse any XSD content model node into a plain dict tree.
    Returns None for non-structural nodes (xsd:attribute, xsd:annotation, etc.).
    """
    tag = local(elem.tag)
    min_o, max_o = parse_occurs(elem)

    if tag == "element":
        name = elem.get("name")
        if not name:
            # ref="prefix:local" — strip prefix to get local name
            ref = elem.get("ref", "")
            name = ref.split(":")[-1] if ref else None
        if name:
            return {"kind": "element", "name": name, "min": min_o, "max": max_o}

    elif tag == "group":
        ref = elem.get("ref", "")
        # Only handle group *references* (ref=) here; definitions (name=) are handled separately.
        gname = ref.split(":")[-1] if ref else None
        if gname:
            return {"kind": "group_ref", "name": gname, "min": min_o, "max": max_o}

    elif tag in ("sequence", "choice", "all"):
        items = [parse_node(child) for child in elem]
        items = [i for i in items if i is not None]
        return {"kind": tag, "min": min_o, "max": max_o, "items": items}

    elif tag == "any":
        return {"kind": "any", "min": min_o, "max": max_o}

    return None


def parse_ct_model(ct_elem) -> dict | None:
    """Extract the content model node from a complexType element."""
    for child in ct_elem:
        tag = local(child.tag)
        if tag in ("sequence", "choice", "all"):
            return parse_node(child)
        elif tag == "complexContent":
            # Extension/restriction: look one level deeper for the compositor
            for cc_child in child:
                cc_tag = local(cc_child.tag)
                if cc_tag in ("extension", "restriction"):
                    for ext_child in cc_child:
                        ext_tag = local(ext_child.tag)
                        if ext_tag in ("sequence", "choice", "all", "group"):
                            return parse_node(ext_child)
    return None


def parse_group_model(group_elem) -> dict | None:
    """Extract the content model node from a group definition element."""
    for child in group_elem:
        tag = local(child.tag)
        if tag in ("sequence", "choice", "all"):
            return parse_node(child)
    return None


# Tags that cannot contain element declarations — no need to recurse into them.
_SKIP_TAGS = frozenset(
    (
        "simpleType",
        "annotation",
        "documentation",
        "restriction",
        "enumeration",
        "pattern",
        "union",
        "list",
        "attribute",
        "attributeGroup",
    )
)


def collect_elem_decls(xml_elem, result: dict) -> None:
    """
    Walk the XML tree and collect every local element declaration of the form
    <xsd:element name="X" type="CT_X">, recording name → CT_name pairs.

    This is necessary because OOXML XSD declares most elements locally inside
    complex types rather than at the schema root level. The same element name
    always maps to the same CT within a given ML type, so last-writer-wins is fine.
    """
    tag = local(xml_elem.tag)
    if tag in _SKIP_TAGS:
        return

    if tag == "element":
        name = xml_elem.get("name")
        type_attr = xml_elem.get("type", "")
        if name and type_attr:
            # Strip namespace prefix from type reference (e.g. "s:ST_Foo" → "ST_Foo")
            result[name] = type_attr.split(":")[-1]

    for child in xml_elem:
        collect_elem_decls(child, result)


def parse_xsd_file(
    xsd_path: Path,
    ml_type: str,
    elem_registry: dict,
    type_registry: dict,
    group_registry: dict,
) -> None:
    """
    Parse one XSD file and populate the three registries:
      elem_registry:  (local_name, ml_type) → ct_name
      type_registry:  (ct_name, ml_type) → content model node
      group_registry: (group_name, ml_type) → content model node
    """
    try:
        tree = ET.parse(xsd_path)
    except ET.ParseError as e:
        print(f"  Warning: failed to parse {xsd_path.name}: {e}", file=sys.stderr)
        return

    root = tree.getroot()

    # Pass 1: collect CT and group content models from global definitions
    for child in root:
        tag = local(child.tag)
        if tag == "complexType":
            ct_name = child.get("name")
            if ct_name:
                model = parse_ct_model(child)
                if model:
                    type_registry[(ct_name, ml_type)] = model
        elif tag == "group":
            gname = child.get("name")
            if gname:
                model = parse_group_model(child)
                if model:
                    group_registry[(gname, ml_type)] = model

    # Pass 2: collect element name → CT type mappings from the entire file.
    # OOXML XSD declares most elements locally inside complex types, not at the
    # schema root, so we must walk the full tree to find all (name, type) pairs.
    local_decls: dict = {}
    collect_elem_decls(root, local_decls)
    for elem_name, ct_name in local_decls.items():
        elem_registry[(elem_name, ml_type)] = ct_name


def fmt_occurs(min_o: int, max_o: str) -> str:
    if min_o == 1 and max_o == "1":
        return ""  # required exactly once: no annotation needed
    max_s = "*" if max_o == "unbounded" else max_o
    return f" [{min_o}..{max_s}]"


def render_node(
    node: dict,
    all_groups: dict,
    indent: int = 0,
    depth: int = 0,
    seen_groups: frozenset = frozenset(),
    prefix_fn=None,
) -> str:
    """Render a content model node as indented human-readable text."""
    if depth > 8:
        return "  " * indent + "..."

    pad = "  " * indent
    kind = node["kind"]

    if kind == "element":
        name = prefix_fn(node["name"]) if prefix_fn else node["name"]
        occ = fmt_occurs(node["min"], node["max"])
        return f"{pad}{name}{occ}"

    elif kind == "any":
        occ = fmt_occurs(node["min"], node["max"])
        return f"{pad}(any){occ}"

    elif kind == "group_ref":
        gname = node["name"]
        if gname in seen_groups:
            return f"{pad}(group {gname} — recursive)"
        group_node = all_groups.get(gname)
        if group_node is None:
            occ = fmt_occurs(node["min"], node["max"])
            return f"{pad}(group {gname}){occ}"
        # Merge the group ref's cardinality onto the inner compositor
        merged = dict(group_node, min=node["min"], max=node["max"])
        return render_node(merged, all_groups, indent, depth + 1, seen_groups | {gname}, prefix_fn)

    elif kind in ("sequence", "choice", "all"):
        occ = fmt_occurs(node["min"], node["max"])
        header = f"{pad}{kind}{occ}:"
        item_lines = []
        for item in node.get("items", []):
            rendered = render_node(item, all_groups, indent + 1, depth + 1, seen_groups, prefix_fn)
            if rendered:
                item_lines.append(rendered)
        if not item_lines:
            return ""
        return "\n".join([header] + item_lines)

    return ""


def collect_element_names(
    node: dict,
    all_groups: dict,
    seen_groups: frozenset = frozenset(),
) -> set[str]:
    """Return the set of all element local_names reachable from a content model node."""
    names: set[str] = set()
    kind = node["kind"]

    if kind == "element":
        names.add(node["name"])
    elif kind == "group_ref":
        gname = node["name"]
        if gname not in seen_groups:
            group_node = all_groups.get(gname)
            if group_node:
                names |= collect_element_names(group_node, all_groups, seen_groups | {gname})
    elif kind in ("sequence", "choice", "all"):
        for item in node.get("items", []):
            names |= collect_element_names(item, all_groups, seen_groups)
    return names


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build schema parent/child index from ECMA-376 XSD files."
    )
    parser.add_argument(
        "--xsd-dir",
        choices=["strict", "transitional"],
        default="transitional",
        help="Which XSD directory to parse (default: transitional)",
    )
    args = parser.parse_args()

    xsd_dir = SOURCE_DIR / f"{args.xsd_dir}-xsd"
    if not xsd_dir.exists():
        print(f"Error: XSD directory not found: {xsd_dir}", file=sys.stderr)
        sys.exit(1)

    if not DB_PATH.exists():
        print(
            f"Error: index.db not found at {DB_PATH}\n"
            "Run build_index.py first to create the database.",
            file=sys.stderr,
        )
        sys.exit(1)

    # Registries populated by parsing
    elem_registry: dict = {}   # (local_name, ml_type) → ct_name
    type_registry: dict = {}   # (ct_name, ml_type) → content model node
    group_registry: dict = {}  # (group_name, ml_type) → content model node

    xsd_files = sorted(xsd_dir.glob("*.xsd"))
    print(f"Parsing {len(xsd_files)} XSD files from {xsd_dir.name}/...")

    for xsd_path in xsd_files:
        ml_type = file_to_ml(xsd_path.name)
        parse_xsd_file(xsd_path, ml_type, elem_registry, type_registry, group_registry)

    print(f"  Elements:      {len(elem_registry)}")
    print(f"  Complex types: {len(type_registry)}")
    print(f"  Groups:        {len(group_registry)}")

    # Flat group lookup for rendering and element collection. Group names don't
    # collide across OOXML XSD files, so a flat dict keyed by name alone is safe.
    all_groups: dict = {gname: node for (gname, _), node in group_registry.items()}

    # Reverse lookup: element name → set of ml_types it's registered in.
    # Used to determine the correct namespace prefix for cross-namespace refs.
    name_to_mls: dict[str, set[str]] = defaultdict(set)
    for (name, mt) in elem_registry:
        name_to_mls[name].add(mt)

    def make_prefix_fn(default_ml: str):
        """Return a function that prepends the namespace prefix to an element name."""
        default_pfx = ML_TO_PREFIX.get(default_ml, "")

        def prefix_fn(name: str) -> str:
            mls = name_to_mls.get(name, set())
            if default_ml in mls:
                pfx = default_pfx
            elif mls:
                pfx = ML_TO_PREFIX.get(next(iter(mls)), "")
            else:
                pfx = default_pfx
            return f"{pfx}:{name}" if pfx else name

        return prefix_fn

    # ── Children ──────────────────────────────────────────────────────────────
    print("Building children data...")
    children_data: dict = {}  # (local_name, ml_type) → rendered content model
    for (elem_name, ml_type), ct_name in elem_registry.items():
        model = type_registry.get((ct_name, ml_type))
        if model is None:
            continue
        rendered = render_node(model, all_groups, prefix_fn=make_prefix_fn(ml_type))
        if rendered:
            children_data[(elem_name, ml_type)] = rendered

    # ── Parents ───────────────────────────────────────────────────────────────
    print("Building parent data...")

    # For each CT, find all element names that can appear in it
    ct_to_child_names: dict = {}
    for (ct_name, ml_type), model in type_registry.items():
        names = collect_element_names(model, all_groups)
        if names:
            ct_to_child_names[(ct_name, ml_type)] = names

    # Invert: child element name → set of (ct_name, ml_type) that contain it
    child_name_to_cts: dict = defaultdict(set)
    for (ct_name, ml_type), names in ct_to_child_names.items():
        for name in names:
            child_name_to_cts[name].add((ct_name, ml_type))

    # Invert elem_registry: (ct_name, ml_type) → set of (elem_name, ml_type) pairs
    ct_key_to_elems: dict[tuple, set[tuple[str, str]]] = defaultdict(set)
    for (elem_name, ml_type), ct_name in elem_registry.items():
        ct_key_to_elems[(ct_name, ml_type)].add((elem_name, ml_type))

    def prefixed_name(name: str, ml: str) -> str:
        pfx = ML_TO_PREFIX.get(ml, "")
        return f"{pfx}:{name}" if pfx else name

    # For each element, collect all parent elements with their ml_type for prefixing
    parents_data: dict = {}  # (elem_name, ml_type) → sorted list of prefixed parent names
    for (elem_name, ml_type) in elem_registry:
        parent_entries: set[tuple[str, str]] = set()
        for ct_key in child_name_to_cts.get(elem_name, set()):
            parent_entries |= ct_key_to_elems.get(ct_key, set())
        if parent_entries:
            parents_data[(elem_name, ml_type)] = sorted(
                prefixed_name(n, mt) for n, mt in parent_entries
            )

    # ── Write to DB ───────────────────────────────────────────────────────────
    print("Writing schema tables to index.db...")
    conn = sqlite3.connect(DB_PATH)
    conn.executescript("""
        DROP TABLE IF EXISTS schema_parents;
        DROP TABLE IF EXISTS schema_children;
        CREATE TABLE schema_parents (
            local_name  TEXT NOT NULL,
            ml_type     TEXT NOT NULL,
            parents     TEXT NOT NULL,
            PRIMARY KEY (local_name, ml_type)
        );
        CREATE TABLE schema_children (
            local_name     TEXT NOT NULL,
            ml_type        TEXT NOT NULL,
            content_model  TEXT NOT NULL,
            PRIMARY KEY (local_name, ml_type)
        );
    """)

    conn.executemany(
        "INSERT INTO schema_parents VALUES (?, ?, ?)",
        [(k[0], k[1], ", ".join(v)) for k, v in parents_data.items()],
    )
    conn.executemany(
        "INSERT INTO schema_children VALUES (?, ?, ?)",
        [(k[0], k[1], v) for k, v in children_data.items()],
    )
    conn.commit()
    conn.close()

    # ── Summary ───────────────────────────────────────────────────────────────
    from collections import Counter

    print("\nSchema index complete.")
    print(f"  Elements with children info: {len(children_data)}")
    print(f"  Elements with parents info:  {len(parents_data)}")
    print()

    ml_child_counts = Counter(ml for (_, ml) in children_data)
    ml_parent_counts = Counter(ml for (_, ml) in parents_data)
    print(f"{'ML type':<35} {'Children':>10} {'Parents':>10}")
    print("-" * 57)
    for ml in sorted(ml_child_counts):
        print(f"{ml:<35} {ml_child_counts[ml]:>10} {ml_parent_counts.get(ml, 0):>10}")


if __name__ == "__main__":
    main()
