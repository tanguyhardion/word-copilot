import re
from word_copilot.dsl.parser import tokenize_dsl_line, extract_fields, parse_dsl

ID_COMMENT_RE = re.compile(r"//\s*id=(\d+)\s*$")


def looks_like_edit_dsl(dsl_string):
    for line in dsl_string.strip().split("\n"):
        line = line.strip()
        if not line or line.startswith("//") or line.startswith("#"):
            continue
        return line.lower().startswith("edit ") or line.lower() == "edit"
    return False


def _read_block_until_endblock(raw_lines, i):
    """
    Starting just after an op line, collect raw lines until a line that is
    exactly 'endblock' (case-insensitive). Returns (block_lines, next_index).
    """
    block_lines = []
    while i < len(raw_lines):
        stripped = raw_lines[i].strip()
        if stripped.lower() == "endblock":
            return block_lines, i + 1
        block_lines.append(raw_lines[i])
        i += 1
    print("[edit-dsl] warning: block not terminated with 'endblock'")
    return block_lines, i


def parse_edit_dsl(dsl_string):
    """
    Parses edit-mode DSL into a list of operation dicts:
      {"op": "delete", "target_id": N}
      {"op": "replace", "target_id": N, "elements": [...]}
      {"op": "insert_before", "target_id": N, "elements": [...]}
      {"op": "insert_after",  "target_id": N, "elements": [...]}
      {"op": "insert_at", "position": "end", "elements": [...]}
    Element sub-blocks are parsed with the existing parse_dsl(), so every
    normal element type (h1-h6, p, ul/ol, table, image, icon, svg, textbox)
    works unchanged inside replace/insert_* blocks.
    """
    raw_lines = dsl_string.strip().split("\n")
    ops = []
    i = 0
    while i < len(raw_lines):
        raw = raw_lines[i]
        line = raw.strip()

        if not line or line.startswith("//") or line.startswith("#"):
            i += 1
            continue

        tokens, _ = tokenize_dsl_line(line)
        if not tokens:
            i += 1
            continue

        op_type, fields = extract_fields(tokens)

        if op_type == "edit":
            # Just a header line, e.g. `edit target=active` — nothing to do.
            i += 1
            continue

        if op_type == "delete":
            if "id" not in fields:
                print(f"[edit-dsl] 'delete' missing id= — skipped")
                i += 1
                continue
            ops.append({"op": "delete", "target_id": int(fields["id"])})
            i += 1
            continue

        if op_type in ("replace", "insert_before", "insert_after"):
            if "id" not in fields:
                print(f"[edit-dsl] '{op_type}' missing id= — skipped")
                i += 1
                # still try to consume a dangling block so parsing doesn't desync
                if i < len(raw_lines):
                    _, i = _read_block_until_endblock(raw_lines, i)
                continue
            target_id = int(fields["id"])
            block_lines, i = _read_block_until_endblock(raw_lines, i + 1)
            elements = parse_dsl("\n".join(block_lines))
            ops.append({"op": op_type, "target_id": target_id, "elements": elements})
            continue

        if op_type == "insert_at":
            position = fields.get("position") or (
                tokens[1].lower() if len(tokens) > 1 and "=" not in tokens[1] else "end"
            )
            # support both `insert_at end` and `insert_at position=end`
            if len(tokens) > 1 and "=" not in tokens[1]:
                position = tokens[1].lower()
            else:
                position = fields.get("position", "end").lower()
            block_lines, i = _read_block_until_endblock(raw_lines, i + 1)
            elements = parse_dsl("\n".join(block_lines))
            ops.append({"op": "insert_at", "position": position, "elements": elements})
            continue

        if op_type == "endblock":
            # Stray endblock with no matching op — ignore.
            i += 1
            continue

        print(f"[edit-dsl] unknown edit op '{op_type}' — skipped")
        i += 1

    return ops
