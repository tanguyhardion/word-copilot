import re
from word_copilot.constants import FLOW_TYPES, HEADING_TYPES
from word_copilot.dsl.builders import (
    build_page_directive,
    build_heading,
    build_paragraph,
    build_list_item,
    build_hr,
    build_image_elem,
    build_icon_elem,
    build_svg_elem,
    build_table_elem,
    build_textbox_elem,
    parse_rich_text,
    parse_text_segment,
)

PAGE_SEPARATOR = re.compile(r"^\s*---\s*$")


def _split_on_pipe(s):
    in_q = False
    for i, ch in enumerate(s):
        if ch == '"':
            in_q = not in_q
        elif ch == "|" and not in_q:
            return s[:i], s[i + 1 :]
    return s, None


def _tokenize_kvs(s):
    tokens = []
    pattern = re.compile(r'(\w[\w\-]*(?:=[^\s"]*|="[^"]*")?)')
    for m in pattern.finditer(s):
        tok = m.group(1).strip()
        if tok:
            tokens.append(tok)
    return tokens


def tokenize_dsl_line(line):
    line = re.sub(r"(?<!\S)//.*$", "", line).strip()
    if not line:
        return [], None
    before, after = _split_on_pipe(line)
    tokens = _tokenize_kvs(before)
    return tokens, after.strip() if after is not None else None


def extract_fields(tokens):
    if not tokens:
        return None, {}
    shape_type = tokens[0].lower()
    fields = {}
    for tok in tokens[1:]:
        if "=" in tok:
            k, _, v = tok.partition("=")
            fields[k.strip().lower()] = v.strip().strip('"')
    return shape_type, fields


def parse_dsl(dsl_string):
    raw_lines = dsl_string.strip().split("\n")
    elements = []
    i = 0
    current_list = None

    def flush_list():
        nonlocal current_list
        if current_list:
            elements.append(current_list)
            current_list = None

    while i < len(raw_lines):
        raw = raw_lines[i]
        line = raw.strip()
        source_line = i + 1

        if not line or line.startswith("//") or line.startswith("#"):
            i += 1
            continue

        tokens, text_part = tokenize_dsl_line(line)
        if not tokens:
            i += 1
            continue

        elem_type, fields = extract_fields(tokens)

        if elem_type not in FLOW_TYPES:
            print(f"[dsl] line {source_line}: unknown element '{elem_type}' — skipped")
            i += 1
            continue

        if elem_type == "item":
            if current_list is None:
                elem = build_list_item(fields, text_part, source_line)
                elements.append(elem)
            else:
                item = build_list_item(fields, text_part, source_line)
                current_list["items"].append(item)
            i += 1
            continue

        if elem_type not in ("ul", "ol"):
            flush_list()

        if elem_type == "page":
            elements.append(build_page_directive(fields))
            i += 1

        elif elem_type in HEADING_TYPES:
            elements.append(build_heading(elem_type, fields, text_part, source_line))
            i += 1

        elif elem_type == "p":
            elements.append(build_paragraph(fields, text_part, source_line))
            i += 1

        elif elem_type in ("ul", "ol"):
            flush_list()
            current_list = {
                "type": elem_type,
                "_source_line": source_line,
                "fields": fields,
                "items": [],
            }
            i += 1

        elif elem_type == "hr":
            elements.append(build_hr(fields, source_line))
            i += 1

        elif elem_type in ("br", "pagebreak"):
            elements.append({"type": elem_type, "_source_line": source_line})
            i += 1

        elif elem_type == "image":
            elem = build_image_elem(fields, source_line)
            if elem:
                elements.append(elem)
            i += 1

        elif elem_type == "icon":
            elem = build_icon_elem(fields, source_line)
            if elem:
                elements.append(elem)
            i += 1

        elif elem_type == "svg":
            svg_lines = []
            i += 1
            while i < len(raw_lines):
                nl = raw_lines[i]
                if nl.strip().lower() == "endsvg":
                    i += 1
                    break
                svg_lines.append(nl)
                i += 1
            svg_markup = "\n".join(svg_lines).strip()
            if svg_markup:
                elem = build_svg_elem(fields, svg_markup, source_line)
                elements.append(elem)

        elif elem_type == "table":
            subsequent = []
            i += 1
            while i < len(raw_lines):
                nl = raw_lines[i].strip()
                if not nl or nl.split()[0] in FLOW_TYPES:
                    break
                subsequent.append(nl)
                i += 1
            elem = build_table_elem(fields, subsequent, source_line)
            if elem:
                elements.append(elem)

        elif elem_type == "textbox":
            elements.append(build_textbox_elem(fields, text_part, source_line))
            i += 1

        else:
            i += 1

    flush_list()
    return elements


def parse_dsl_pages(dsl_string):
    raw_lines = dsl_string.strip().split("\n")
    blocks, current = [], []
    for line in raw_lines:
        if PAGE_SEPARATOR.match(line):
            blocks.append("\n".join(current))
            current = []
        else:
            current.append(line)
    if current:
        blocks.append("\n".join(current))
    result = [parse_dsl(b) for b in blocks if b.strip()]
    return result if result else [[]]
