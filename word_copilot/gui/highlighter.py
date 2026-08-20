import re
from word_copilot.constants import FLOW_TYPES, EDIT_TYPES, DSL_COLOR_ALIASES

HL_ELEM_NAMES = sorted(FLOW_TYPES | EDIT_TYPES, key=len, reverse=True)

HL_FIELD_KEYS = [
    "size",
    "orientation",
    "margin",
    "left",
    "top",
    "width",
    "height",
    "rotation",
    "color",
    "transparency",
    "outline",
    "shadow",
    "align",
    "valign",
    "halign",
    "font",
    "bold",
    "italic",
    "underline",
    "spacing_after",
    "spacing_before",
    "line_height",
    "indent",
    "padding",
    "header_fill",
    "header_text_color",
    "header_bold",
    "row_fill",
    "alt_row_fill",
    "text_color",
    "border_color",
    "border_weight",
    "font_size",
    "url",
    "name",
    "style",
    "fit",
    "anchor",
    "bullet_color",
    "size",
    "bold",
    "italic",
    "underline",
    "weight",
    "dash",
    "id",
    "target",
    "position",
]

HL_COLOR_ALIASES = list(DSL_COLOR_ALIASES.keys())


class DSLHighlighter:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self._setup_tags()

    def _setup_tags(self):
        self.text_widget.tag_configure("hl_shape", foreground="#5B9BD5")
        self.text_widget.tag_configure("hl_key", foreground="#9E7CC1")
        self.text_widget.tag_configure("hl_value", foreground="#ED7D31")
        self.text_widget.tag_configure("hl_color", foreground="#70AD47")
        self.text_widget.tag_configure("hl_string", foreground="#FFD966")
        self.text_widget.tag_configure(
            "hl_comment", foreground="#7F7F7F", font=("Consolas", 10, "italic")
        )
        self.text_widget.tag_configure("hl_pipe", foreground="#A5A5A5")
        self.text_widget.tag_configure("hl_equals", foreground="#808080")
        self.text_widget.tag_configure("hl_svg_tag", foreground="#5B9BD5")
        self.text_widget.tag_configure("hl_svg_attr", foreground="#9E7CC1")
        self.text_widget.tag_configure("hl_svg_comment", foreground="#7F7F7F")
        self.text_widget.tag_configure("hl_svg_endsvg", foreground="#ED7D31")
        self.text_widget.tag_configure(
            "hl_separator", foreground="#FFC000", font=("Consolas", 10, "bold")
        )
        self.text_widget.tag_configure(
            "hl_edit_op", foreground="#E06C75", font=("Consolas", 10, "bold")
        )

    def highlight(self, event=None):
        shape_pat = (
            r"^("
            + "|".join(re.escape(s) for s in sorted(FLOW_TYPES, key=len, reverse=True))
            + r")(?=\s|$)"
        )
        edit_op_pat = (
            r"^("
            + "|".join(re.escape(s) for s in sorted(EDIT_TYPES, key=len, reverse=True))
            + r")(?=\s|$)"
        )
        key_pat = (
            r"(?<!\w)(" + "|".join(re.escape(k) for k in HL_FIELD_KEYS) + r")(?==)"
        )
        alias_pat = (
            r"(?<==)("
            + "|".join(re.escape(c) for c in HL_COLOR_ALIASES)
            + r")(?=\s|$|,)"
        )
        variant_pat = r"(?<==)(a[1-6]|bg[12]|t[12])_(l1|l2|d1|d2)(?=\s|$|,)"
        hex_pat = r"#[0-9A-Fa-f]{3,6}\b"
        num_pat = r"(?<==)-?\d+(\.\d+)?"
        str_pat = r'"[^"]*"'
        pipe_pat = r"\|"
        eq_pat = r"="
        svg_tag_pat = r"</?[A-Za-z_:][\w:.\-]*"
        svg_attr_pat = r"(?<=\s)([A-Za-z_:][\w:.\-]*)(?=\=)"
        svg_cmt_pat = r"<!--.*?-->"
        endsvg_pat = r"^\s*endsvg\s*$"

        all_tags = (
            "hl_shape",
            "hl_key",
            "hl_value",
            "hl_color",
            "hl_string",
            "hl_comment",
            "hl_pipe",
            "hl_equals",
            "hl_svg_tag",
            "hl_svg_attr",
            "hl_svg_comment",
            "hl_svg_endsvg",
            "hl_separator",
            "hl_edit_op",
        )
        for tag in all_tags:
            self.text_widget.tag_remove(tag, "1.0", "end")

        content = self.text_widget.get("1.0", "end-1c")
        in_svg = False

        for li, line in enumerate(content.split("\n")):
            row = li + 1

            def py_to_tk(offset, _l=line):
                col = 0
                for i in range(min(offset, len(_l))):
                    col += 2 if ord(_l[i]) > 0xFFFF else 1
                return col

            def add_tag(pat, tag, _l=line, _r=row, flags=0):
                for m in re.finditer(pat, _l, flags):
                    s = m.start(1) if m.lastindex else m.start()
                    e = m.end(1) if m.lastindex else m.end()
                    self.text_widget.tag_add(
                        tag, f"{_r}.{py_to_tk(s,_l)}", f"{_r}.{py_to_tk(e,_l)}"
                    )

            if re.match(r"^\s*endblock\s*$", line, re.IGNORECASE):
                self.text_widget.tag_add("hl_edit_op", f"{row}.0", f"{row}.end")
                continue

            if re.match(edit_op_pat, line):
                for p, t in [
                    (str_pat, "hl_string"),
                    (pipe_pat, "hl_pipe"),
                    (edit_op_pat, "hl_edit_op"),
                    (key_pat, "hl_key"),
                    (num_pat, "hl_value"),
                    (eq_pat, "hl_equals"),
                ]:
                    add_tag(p, t)
                continue

            if re.match(r"^\s*---\s*$", line):
                self.text_widget.tag_add("hl_separator", f"{row}.0", f"{row}.end")
                continue

            if not in_svg and (re.match(r"\s*//", line) or re.match(r"\s*#", line)):
                self.text_widget.tag_add("hl_comment", f"{row}.0", f"{row}.end")
                continue

            if not in_svg and re.match(r"\s*svg(\s|$)", line):
                for p, t in [
                    (str_pat, "hl_string"),
                    (pipe_pat, "hl_pipe"),
                    (shape_pat, "hl_shape"),
                    (alias_pat, "hl_color"),
                    (variant_pat, "hl_color"),
                    (hex_pat, "hl_color"),
                    (key_pat, "hl_key"),
                    (num_pat, "hl_value"),
                    (eq_pat, "hl_equals"),
                ]:
                    add_tag(p, t)
                in_svg = True
                continue

            if in_svg:
                if re.match(endsvg_pat, line):
                    self.text_widget.tag_add("hl_svg_endsvg", f"{row}.0", f"{row}.end")
                    in_svg = False
                    continue
                for p, t in [
                    (svg_cmt_pat, "hl_svg_comment"),
                    (svg_tag_pat, "hl_svg_tag"),
                    (svg_attr_pat, "hl_svg_attr"),
                    (str_pat, "hl_string"),
                    (hex_pat, "hl_color"),
                    (eq_pat, "hl_equals"),
                ]:
                    add_tag(p, t)
                continue

            if re.match(r"\s*(cols|header|row)=", line):
                for p, t in [
                    (r"^(cols|header|row)(?==)", "hl_key"),
                    (str_pat, "hl_string"),
                    (num_pat, "hl_value"),
                    (eq_pat, "hl_equals"),
                ]:
                    add_tag(p, t)
                continue

            for p, t in [
                (str_pat, "hl_string"),
                (pipe_pat, "hl_pipe"),
                (shape_pat, "hl_shape"),
                (alias_pat, "hl_color"),
                (variant_pat, "hl_color"),
                (hex_pat, "hl_color"),
                (key_pat, "hl_key"),
                (num_pat, "hl_value"),
                (eq_pat, "hl_equals"),
            ]:
                add_tag(p, t)
