"""
Build the schema parent/child index from ECMA-376 XSD files.

Writes to tables in index.db: schema_parents and schema_children.
"""

import re
import sqlite3
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

from _prefix_map import PREFIX_MAP

DB_PATH = Path(__file__).parent / "index.db"
SCHEMA_DIR = Path(__file__).parent.parent / "schemas"

# Reverse map: ml_type → namespace prefix (first match wins).
# Used by prefixed_name() when rendering parent element names.
ML_TO_PREFIX: dict[str, str] = {}
for _pfx, (_ml, _uri) in PREFIX_MAP.items():
    if _ml not in ML_TO_PREFIX:
        ML_TO_PREFIX[_ml] = _pfx

# Forward maps derived from PREFIX_MAP: namespace URI → display prefix / ml_type.
URI_TO_PREFIX: dict[str, str] = {}
URI_TO_ML: dict[str, str] = {}
for _pfx, (_ml, _uri) in PREFIX_MAP.items():
    URI_TO_PREFIX.setdefault(_uri, _pfx)
    URI_TO_ML.setdefault(_uri, _ml)

# DrawingML sub-namespaces share the "a" display prefix and "DrawingML" ml_type.
# Their targetNamespace URIs differ from dml-main.xsd, so we add explicit entries.
_DML_SUB_NS = [
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
]
for _uri in _DML_SUB_NS:
    URI_TO_PREFIX.setdefault(_uri, "a")
    URI_TO_ML.setdefault(_uri, "DrawingML")


def extract_ns_info(xsd_path: Path) -> tuple[str, dict[str, str]]:
    """
    Extract targetNamespace and xmlns prefix→URI map from the XSD root element.
    ElementTree strips xmlns declarations, so we use regex on the raw header.
    """
    header = xsd_path.read_text(encoding="utf-8")[:4096]
    target_m = re.search(r'targetNamespace="([^"]+)"', header)
    target_ns = target_m.group(1) if target_m else ""
    ns_map = {
        m.group(1): m.group(2)
        for m in re.finditer(r'xmlns:(\w+)="([^"]+)"', header)
    }
    return target_ns, ns_map


def local(tag: str) -> str:
    """Strip the XSD namespace URI from a qualified ElementTree tag."""
    return tag.split("}")[1] if "}" in tag else tag


def parse_occurs(elem) -> tuple[int, str]:
    min_o = int(elem.get("minOccurs", "1"))
    max_o = elem.get("maxOccurs", "1")
    return min_o, max_o


def parse_node(
    elem, ns_map: dict[str, str], target_ns: str
) -> dict | None:
    """
    Recursively parse any XSD content model node into a plain dict tree.
    Returns None for non-structural nodes (xsd:attribute, xsd:annotation, etc.).
    """
    tag = local(elem.tag)
    min_o, max_o = parse_occurs(elem)

    if tag == "element":
        name = elem.get("name")
        if name:
            # Local declaration — belongs to this file's namespace
            prefix = URI_TO_PREFIX.get(target_ns, "")
        else:
            # ref="prefix:local" — resolve prefix to canonical display prefix
            ref = elem.get("ref", "")
            if ":" in ref:
                ref_pfx, name = ref.split(":", 1)
                ref_uri = ns_map.get(ref_pfx, "")
                prefix = URI_TO_PREFIX.get(ref_uri, ref_pfx)
            else:
                name = ref or None
                prefix = URI_TO_PREFIX.get(target_ns, "")
        if name:
            return {
                "kind": "element", "name": name, "prefix": prefix,
                "min": min_o, "max": max_o,
            }

    elif tag == "group":
        ref = elem.get("ref", "")
        # Only handle group *references* (ref=) here; definitions (name=) are handled separately.
        gname = ref.split(":")[-1] if ref else None
        if gname:
            return {"kind": "group_ref", "name": gname, "min": min_o, "max": max_o}

    elif tag in ("sequence", "choice", "all"):
        items = [parse_node(child, ns_map, target_ns) for child in elem]
        items = [i for i in items if i is not None]
        return {"kind": tag, "min": min_o, "max": max_o, "items": items}

    elif tag == "any":
        return {"kind": "any", "min": min_o, "max": max_o}

    return None


def parse_ct_model(
    ct_elem, ns_map: dict[str, str], target_ns: str
) -> dict | None:
    """Extract the content model node from a complexType element."""
    for child in ct_elem:
        tag = local(child.tag)
        if tag in ("sequence", "choice", "all"):
            return parse_node(child, ns_map, target_ns)
        elif tag == "complexContent":
            # Extension/restriction: look one level deeper for the compositor
            for cc_child in child:
                cc_tag = local(cc_child.tag)
                if cc_tag in ("extension", "restriction"):
                    for ext_child in cc_child:
                        ext_tag = local(ext_child.tag)
                        if ext_tag in ("sequence", "choice", "all", "group"):
                            return parse_node(ext_child, ns_map, target_ns)
    return None


def parse_group_model(
    group_elem, ns_map: dict[str, str], target_ns: str
) -> dict | None:
    """Extract the content model node from a group definition element."""
    for child in group_elem:
        tag = local(child.tag)
        if tag in ("sequence", "choice", "all"):
            return parse_node(child, ns_map, target_ns)
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
    target_ns, ns_map = extract_ns_info(xsd_path)
    ml_type = URI_TO_ML.get(target_ns, "Shared")

    tree = ET.parse(xsd_path)
    root = tree.getroot()

    # Pass 1: collect CT and group content models from global definitions
    for child in root:
        tag = local(child.tag)
        if tag == "complexType":
            ct_name = child.get("name")
            if ct_name:
                model = parse_ct_model(child, ns_map, target_ns)
                if model:
                    type_registry[(ct_name, ml_type)] = model
        elif tag == "group":
            gname = child.get("name")
            if gname:
                model = parse_group_model(child, ns_map, target_ns)
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
    max_s = "*" if max_o == "unbounded" else max_o
    return f" [{min_o}..{max_s}]"


def render_node(
    node: dict,
    all_groups: dict,
    indent: int = 0,
    depth: int = 0,
    seen_groups: frozenset = frozenset(),
) -> str:
    """Render a content model node as indented human-readable text."""
    if depth > 8:
        return "  " * indent + "..."

    pad = "  " * indent
    kind = node["kind"]

    if kind == "element":
        pfx = node.get("prefix", "")
        name = f"{pfx}:{node['name']}" if pfx else node["name"]
        occ = fmt_occurs(node["min"], node["max"])
        return f"{pad}{name}{occ}"

    elif kind == "any":
        occ = fmt_occurs(node["min"], node["max"])
        return f"{pad}(any){occ}"

    elif kind == "group_ref":
        gname = node["name"]
        if gname in seen_groups:
            return f"{pad}(group {gname} - recursive)"  # Does not actually occur in current XSDs
        group_node = all_groups.get(gname)
        if group_node is None:
            occ = fmt_occurs(node["min"], node["max"])
            return f"{pad}(group {gname}){occ}"
        # Merge the group ref's cardinality onto the inner compositor
        merged = dict(group_node, min=node["min"], max=node["max"])
        return render_node(merged, all_groups, indent, depth + 1, seen_groups | {gname})

    elif kind in ("sequence", "choice", "all"):
        occ = fmt_occurs(node["min"], node["max"])
        header = f"{pad}{kind}{occ}:"
        item_lines = []
        for item in node.get("items", []):
            rendered = render_node(item, all_groups, indent + 1, depth + 1, seen_groups)
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


def prefixed_name(name: str, ml: str) -> str:
    """Prepend the canonical namespace prefix for an ML type to an element name."""
    pfx = ML_TO_PREFIX.get(ml, "")
    return f"{pfx}:{name}" if pfx else name


def build_children_data(
    elem_registry: dict,
    type_registry: dict,
    all_groups: dict,
) -> dict:
    """For each element, render its content model as human-readable text."""
    children_data: dict = {}
    for (elem_name, ml_type), ct_name in elem_registry.items():
        model = type_registry.get((ct_name, ml_type))
        if model is None:
            continue
        rendered = render_node(model, all_groups)
        if rendered:
            children_data[(elem_name, ml_type)] = rendered
    return children_data


def build_parents_data(
    elem_registry: dict,
    type_registry: dict,
    all_groups: dict,
) -> dict:
    """For each element, find all parent elements that can contain it."""
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

    parents_data: dict = {}
    for (elem_name, ml_type) in elem_registry:
        parent_entries: set[tuple[str, str]] = set()
        for ct_key in child_name_to_cts.get(elem_name, set()):
            parent_entries |= ct_key_to_elems.get(ct_key, set())
        if parent_entries:
            parents_data[(elem_name, ml_type)] = sorted(
                prefixed_name(n, mt) for n, mt in parent_entries
            )
    return parents_data


def init_db(path: Path):
    conn = sqlite3.connect(path)

    conn.executescript("""
        DROP TABLE IF EXISTS schema_parents;
        CREATE TABLE schema_parents (
            local_name  TEXT NOT NULL,
            ml_type     TEXT NOT NULL,
            parents     TEXT NOT NULL,
            PRIMARY KEY (local_name, ml_type)
        );

        DROP TABLE IF EXISTS schema_children;
        CREATE TABLE schema_children (
            local_name     TEXT NOT NULL,
            ml_type        TEXT NOT NULL,
            content_model  TEXT NOT NULL,
            PRIMARY KEY (local_name, ml_type)
        );
    """)

    return conn


def main() -> None:
    if not SCHEMA_DIR.exists():
        raise FileNotFoundError(f"Schema directory not found: {SCHEMA_DIR}")

    # Registries populated by parsing
    elem_registry: dict = {}   # (local_name, ml_type) → ct_name
    type_registry: dict = {}   # (ct_name, ml_type) → content model node
    group_registry: dict = {}  # (group_name, ml_type) → content model node

    xsd_files = sorted(SCHEMA_DIR.glob("*.xsd"))
    print(f"Parsing {len(xsd_files)} XSD files from {SCHEMA_DIR.name}/...")

    for xsd_path in xsd_files:
        parse_xsd_file(xsd_path, elem_registry, type_registry, group_registry)

    print(f"  Elements:      {len(elem_registry)}")
    print(f"  Complex types: {len(type_registry)}")
    print(f"  Groups:        {len(group_registry)}")

    # Flat group lookup for rendering and element collection. Group names don't
    # collide across OOXML XSD files, so a flat dict keyed by name alone is safe.
    all_groups: dict = {gname: node for (gname, _), node in group_registry.items()}

    print("Building children data... ", end="", flush=True)
    children_data = build_children_data(elem_registry, type_registry, all_groups)
    print("done")

    print("Building parent data... ", end="", flush=True)
    parents_data = build_parents_data(elem_registry, type_registry, all_groups)
    print("done")

    print("Writing tables to index.db... ", end="", flush=True)
    conn = init_db(DB_PATH)
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
    print("done")

    from collections import Counter

    print("\nSchema index complete.")
    ml_child_counts = Counter(ml for (_, ml) in children_data)
    ml_parent_counts = Counter(ml for (_, ml) in parents_data)
    for ml in sorted(ml_child_counts):
        c = ml_child_counts[ml]
        p = ml_parent_counts.get(ml, 0)
        print(f"{ml:<30} {c:>5} children {p:>5} parents")
