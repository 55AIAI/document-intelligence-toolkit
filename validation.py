"""Validate document trees and conservative literal support of extracted facts."""
import re


def validate_document(data):
    if not isinstance(data, dict):
        raise ValueError("Document must be a JSON object")
    if not isinstance(data.get("contractors"), list):
        raise ValueError("contractors must be a list")
    active = set()
    count = 0

    def node(value, depth):
        nonlocal count
        if not isinstance(value, dict):
            raise ValueError("Each organization must be an object")
        if id(value) in active or depth > 32:
            raise ValueError("Organization tree contains a cycle or exceeds 32 levels")
        count += 1
        if count > 1000:
            raise ValueError("Document exceeds 1000 organization occurrences")
        active.add(id(value))
        for key in ("name", "bin", "amount", "year", "description", "activity", "address"):
            if value.get(key) is not None and not isinstance(value[key], str):
                raise ValueError(f"Organization field {key} must be text or null")
        for key in ("director", "header", "contract"):
            if value.get(key) is not None:
                if not isinstance(value[key], dict):
                    raise ValueError(f"{key} must be an object or null")
                if any(v is not None and not isinstance(v, str) for v in value[key].values()):
                    raise ValueError(f"{key} fields must be text or null")
        violations = value.get("violations") or []
        if not isinstance(violations, list) or any(not isinstance(v, str) for v in violations):
            raise ValueError("violations must be a list of strings")
        children = value.get("subcontractors") or []
        if not isinstance(children, list):
            raise ValueError("subcontractors must be a list")
        for child in children:
            node(child, depth + 1)
        active.remove(id(value))

    if data.get("customer") is not None:
        node(data["customer"], 0)
    for contractor in data["contractors"]:
        node(contractor, 1)
    return data


def unsupported_facts(data, source):
    """Check literal support; this does not prove the correct relationship."""
    def normalize(value):
        return re.sub(r"[\W_]+", "", value.casefold())

    haystack = normalize(source)
    missing = []

    def walk(org):
        for key in ("name", "bin", "amount", "year"):
            value = org.get(key)
            if value and normalize(value) not in haystack:
                missing.append(key)
        for group, fields in (("director", ("fio", "iin", "phone")),
                              ("contract", ("amount", "year"))):
            for key in fields:
                value = (org.get(group) or {}).get(key)
                if value and normalize(value) not in haystack:
                    missing.append(f"{group}.{key}")
        for child in org.get("subcontractors") or []:
            walk(child)

    if data.get("customer"):
        walk(data["customer"])
    for org in data["contractors"]:
        walk(org)
    return missing


def validate_editor(editor):
    """Check saved canvas identifiers before interpolating them in DOM selectors."""
    if not isinstance(editor, dict):
        raise ValueError("Saved editor must be an object")
    blocks, edges = editor.get("blocks"), editor.get("edges")
    if not isinstance(blocks, list) or not isinstance(edges, list):
        raise ValueError("Saved editor requires blocks and edges lists")
    if len(blocks) > 1000 or len(edges) > 3000:
        raise ValueError("Saved editor is too large")
    ids = set()
    for block in blocks:
        if not isinstance(block, dict) or not re.fullmatch(r"[a-zA-Z0-9_-]{1,80}", str(block.get("id", ""))):
            raise ValueError("Invalid editor block ID")
        if block["id"] in ids:
            raise ValueError("Duplicate editor block ID")
        ids.add(block["id"])
        if block.get("type") not in {"org", "contract", "photo", "text", "shape"}:
            raise ValueError("Unknown editor block type")
        if not re.fullmatch(r"[a-z_-]{1,30}", str(block.get("role", ""))):
            raise ValueError("Invalid editor block role")
        if block["type"] in {"contract", "text", "shape"} and not isinstance(block.get("fields"), dict):
            raise ValueError("Text/contract blocks require fields")
        if block["type"] == "org" and not isinstance(block.get("lines"), list):
            raise ValueError("Organization blocks require lines")
    for edge in edges:
        if not isinstance(edge, dict) or edge.get("from") not in ids or edge.get("to") not in ids:
            raise ValueError("Editor edge references an unknown block")
        if not re.fullmatch(r"[a-zA-Z0-9_-]{1,80}", str(edge.get("id", ""))):
            raise ValueError("Invalid editor edge ID")
    return editor
