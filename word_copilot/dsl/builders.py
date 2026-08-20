import re
from word_copilot.colors import dsl_resolve_color
from word_copilot.constants import MARGIN_PX, VALID_ICON_STYLES


def parse_text_segment(seg):
    seg = seg.strip()
    m = re.match(r'"((?:[^"\\]|\\.)*)"(.*)', seg)
    if not m:
        return None
    text = m.group(1).replace("\\n", "\n").replace('\\"', '"')
    rest = m.group(2).strip()
    result = {"text": text}
    for tok in rest.split():
        if "=" not in tok:
            continue
        k, _, v = tok.partition("=")
        k, v = k.strip().lower(), v.strip()
        if k == "size":
            try:
                result["size"] = int(v)
            except ValueError:
                pass
        elif k == "bold":
            result["bold"] = v.lower() == "true"
        elif k == "italic":
            result["italic"] = v.lower() == "true"
        elif k == "underline":
            result["underline"] = v.lower() == "true"
        elif k == "color":
            resolved = dsl_resolve_color(v)
            if resolved:
                result["color"] = resolved
        elif k == "font":
            result["font"] = v
    return result


def parse_rich_text(text_part):
    segments, current, in_q = [], "", False
    for ch in text_part:
        if ch == '"':
            in_q = not in_q
            current += ch
        elif ch == "+" and not in_q:
            segments.append(current.strip())
            current = ""
        else:
            current += ch
    if current.strip():
        segments.append(current.strip())
    return [r for r in (parse_text_segment(s) for s in segments) if r]


def build_page_directive(fields):
    d = {"type": "page"}
    size = fields.get("size", "letter").lower()
    d["size"] = size
    d["orientation"] = fields.get("orientation", "portrait").lower()
    if "margin" in fields:
        parts = fields["margin"].split(",")
        try:
            vals = [float(p.strip()) for p in parts]
            d["margin_top"] = vals[0] if len(vals) > 0 else MARGIN_PX
            d["margin_right"] = vals[1] if len(vals) > 1 else MARGIN_PX
            d["margin_bottom"] = vals[2] if len(vals) > 2 else MARGIN_PX
            d["margin_left"] = vals[3] if len(vals) > 3 else MARGIN_PX
        except ValueError:
            pass
    return d


def build_heading(elem_type, fields, text_part, source_line=None):
    elem = {
        "type": elem_type,
        "_source_line": source_line,
        "level": int(elem_type[1]),
    }
    if text_part:
        elem["rich_text"] = parse_rich_text(text_part)
    if "size" in fields:
        try:
            elem["size"] = int(fields["size"])
        except ValueError:
            pass
    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            elem["color"] = resolved
    if "bold" in fields:
        elem["bold"] = fields["bold"].lower() == "true"
    if "font" in fields:
        elem["font"] = fields["font"]
    if "align" in fields:
        a = fields["align"].lower()
        if a in ("left", "center", "right", "justify"):
            elem["align"] = a
    if "spacing_after" in fields:
        try:
            elem["spacing_after"] = float(fields["spacing_after"])
        except ValueError:
            pass
    if "spacing_before" in fields:
        try:
            elem["spacing_before"] = float(fields["spacing_before"])
        except ValueError:
            pass
    return elem


def build_paragraph(fields, text_part, source_line=None):
    elem = {"type": "p", "_source_line": source_line}
    if text_part:
        elem["rich_text"] = parse_rich_text(text_part)
    if "size" in fields:
        try:
            elem["size"] = int(fields["size"])
        except ValueError:
            pass
    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            elem["color"] = resolved
    if "bold" in fields:
        elem["bold"] = fields["bold"].lower() == "true"
    if "italic" in fields:
        elem["italic"] = fields["italic"].lower() == "true"
    if "font" in fields:
        elem["font"] = fields["font"]
    if "align" in fields:
        a = fields["align"].lower()
        if a in ("left", "center", "right", "justify"):
            elem["align"] = a
    for sp in ("spacing_after", "spacing_before", "line_height"):
        if sp in fields:
            try:
                elem[sp] = float(fields[sp])
            except ValueError:
                pass
    return elem


def build_list_item(fields, text_part, source_line=None):
    elem = {"type": "item", "_source_line": source_line}
    if text_part:
        elem["rich_text"] = parse_rich_text(text_part)
    if "size" in fields:
        try:
            elem["size"] = int(fields["size"])
        except ValueError:
            pass
    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            elem["color"] = resolved
    if "bold" in fields:
        elem["bold"] = fields["bold"].lower() == "true"
    if "italic" in fields:
        elem["italic"] = fields["italic"].lower() == "true"
    if "font" in fields:
        elem["font"] = fields["font"]
    return elem


def build_hr(fields, source_line=None):
    elem = {"type": "hr", "_source_line": source_line}
    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            elem["color"] = resolved
    if "weight" in fields:
        try:
            elem["weight"] = float(fields["weight"])
        except ValueError:
            pass
    return elem


def build_image_elem(fields, source_line=None):
    url = fields.get("url", "").strip()
    if not url:
        print(f"[dsl] line {source_line}: image missing 'url='")
        return None
    elem = {
        "type": "image",
        "url": url,
        "_source_line": source_line,
    }
    if "width" in fields:
        try:
            elem["width"] = float(fields["width"])
        except ValueError:
            pass
    if "height" in fields:
        try:
            elem["height"] = float(fields["height"])
        except ValueError:
            pass
    align = fields.get("align", "left").lower()
    elem["align"] = align if align in ("left", "center", "right") else "left"
    return elem


def build_icon_elem(fields, source_line=None):
    name = fields.get("name", "").strip()
    if not name:
        print(f"[dsl] line {source_line}: icon missing 'name='")
        return None
    style = fields.get("style", "solid").lower()
    elem = {
        "type": "icon",
        "icon_name": name,
        "icon_style": style if style in VALID_ICON_STYLES else "solid",
        "_source_line": source_line,
    }
    if "width" in fields:
        try:
            elem["width"] = float(fields["width"])
        except ValueError:
            pass
    if "height" in fields:
        try:
            elem["height"] = float(fields["height"])
        except ValueError:
            pass
    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            elem["icon_color"] = resolved
    align = fields.get("align", "left").lower()
    elem["align"] = align if align in ("left", "center", "right") else "left"
    return elem


def build_svg_elem(fields, svg_markup, source_line=None):
    elem = {
        "type": "svg",
        "svg_markup": svg_markup,
        "_source_line": source_line,
    }
    for f in ("width", "height"):
        if f in fields:
            try:
                elem[f] = float(fields[f])
            except ValueError:
                pass
    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            elem["svg_color"] = resolved
    align = fields.get("align", "left").lower()
    elem["align"] = align if align in ("left", "center", "right") else "left"
    return elem


def build_table_elem(fields, subsequent_lines, source_line=None):
    elem = {"type": "table", "_source_line": source_line}
    raw_width = fields.get("width", "100%")
    elem["width_spec"] = raw_width

    style = {}
    for key in (
        "header_fill",
        "header_text_color",
        "row_fill",
        "alt_row_fill",
        "text_color",
        "border_color",
    ):
        if key in fields:
            resolved = dsl_resolve_color(fields[key])
            if resolved:
                style[key] = resolved
    if "border_weight" in fields:
        try:
            style["border_weight"] = float(fields["border_weight"])
        except ValueError:
            pass
    if "font" in fields:
        style["font"] = fields["font"]
    if "font_size" in fields:
        try:
            style["font_size"] = float(fields["font_size"])
        except ValueError:
            pass
    if "header_bold" in fields:
        style["header_bold"] = fields["header_bold"].lower() == "true"
    if "align" in fields:
        a = fields["align"].lower()
        if a in ("left", "center", "right"):
            style["align"] = a
    elem["table_style"] = style

    table = {"content": [], "header_row": False, "col_widths_spec": []}
    for line in subsequent_lines:
        line = line.strip()
        if line.startswith("cols="):
            table["col_widths_spec"] = [w.strip() for w in line[5:].split(",")]
        elif line.startswith("header="):
            cells = re.findall(r'"([^"]*)"', line)
            if cells:
                table["content"].insert(0, cells)
                table["header_row"] = True
        elif line.startswith("row="):
            cells = re.findall(r'"([^"]*)"', line)
            if cells:
                table["content"].append(cells)
    table["rows"] = len(table["content"])
    table["cols"] = max((len(r) for r in table["content"]), default=0)
    elem["table"] = table
    return elem


def build_textbox_elem(fields, text_part, source_line=None):
    elem = {
        "type": "textbox",
        "_source_line": source_line,
        "anchor": fields.get("anchor", "page").lower(),
    }
    for f in ("left", "top", "width", "height"):
        if f in fields:
            try:
                elem[f] = float(fields[f])
            except ValueError:
                pass
    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            elem["color"] = resolved
    if "outline" in fields:
        parts = fields["outline"].split(",")
        color = dsl_resolve_color(parts[0].strip())
        weight = 1.0
        if len(parts) > 1:
            try:
                weight = float(parts[1].strip())
            except ValueError:
                pass
        if color:
            elem["outline"] = {"color": color, "weight": weight}
    if "rotation" in fields:
        try:
            elem["rotation"] = float(fields["rotation"])
        except ValueError:
            pass
    if "transparency" in fields:
        try:
            elem["transparency"] = float(fields["transparency"])
        except ValueError:
            pass
    if text_part:
        elem["rich_text"] = parse_rich_text(text_part)
    for k in ("valign", "halign"):
        if k in fields:
            elem[k] = fields[k].lower()
    if "padding" in fields:
        parts = fields["padding"].split(",")
        try:
            vals = [float(p.strip()) for p in parts]
            elem["padding_left"] = vals[0] if len(vals) > 0 else 0
            elem["padding_right"] = vals[1] if len(vals) > 1 else 0
            elem["padding_top"] = vals[2] if len(vals) > 2 else 0
            elem["padding_bottom"] = vals[3] if len(vals) > 3 else 0
        except ValueError:
            pass
    return elem
