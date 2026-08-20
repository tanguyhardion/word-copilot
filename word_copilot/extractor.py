import os
import re
import tempfile
from word_copilot.colors import THEME_COLORS, DSL_COLOR_ALIASES
from word_copilot.constants import INCH_TO_EMU, HEADING_SIZES
from word_copilot.units import emu_to_px, emu_to_pt


class WordExtractor:
    """
    Reads the active Word document (via python-docx on the saved file path,
    obtained through win32com) and converts it to DSL text.

    Extraction coverage:
      ✓ page setup (size, orientation, margins)
      ✓ headings h1–h6 (text, alignment, color, font, size)
      ✓ paragraphs (text, alignment, color, font, size, bold, italic, spacing)
      ✓ bullet lists (ul) and numbered lists (ol) with run-level formatting
      ✓ horizontal rules (bottom-border paragraphs)
      ✓ page breaks
      ✓ blank paragraphs → br
      ✓ tables (cols, header row detection, cell text)
      ✓ inline images → image placeholder comment
      ✓ drawing/textbox shapes → textbox placeholder

    Each top-level body element (paragraph or table, in document order) is
    tagged with a trailing `// id=N` comment. N is simply that element's
    0-based index among body children whose tag is <w:p> or <w:tbl> — the
    exact same rule WordEditor uses to resolve `id=` targets, so IDs stay
    valid as long as you re-extract before editing.
    """

    # ── public entry point ────────────────────────────────────────────────────
    def extract(self, status_cb=None):
        """
        Returns a DSL string representing the active Word document.
        Raises on failure.
        """
        import pythoncom

        pythoncom.CoInitialize()
        try:
            path = self._get_active_doc_path(status_cb)
            if status_cb:
                status_cb("Reading document structure…")
            dsl = self._docx_to_dsl(path, status_cb)
            return dsl
        finally:
            pythoncom.CoUninitialize()

    def _get_active_doc_path(self, status_cb=None):
        import win32com.client, time

        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            raise Exception(
                "Word is not running. Please open a document in Word first."
            )
        if word.Documents.Count == 0:
            raise Exception("No document is open in Word. Please open one first.")
        doc = word.ActiveDocument
        if status_cb:
            status_cb("Snapshotting document…")
        tmp = os.path.join(
            tempfile.gettempdir(), f"word_dsl_extract_{int(time.time()*1000)}.docx"
        )
        try:
            doc.Range().Copy()
            new_doc = word.Documents.Add()
            new_doc.Range().Paste()
            new_doc.SaveAs2(tmp, 16)  # positional: FileName, FileFormat
            new_doc.Close(SaveChanges=False)
        except Exception as e:
            raise Exception(f"Snapshot via copy-doc failed: {e}")
        return tmp

    def _get_active_doc_path_for_inplace_edit(self, status_cb=None):
        import win32com.client

        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            raise Exception(
                "Word is not running. Please open a document in Word first."
            )
        if word.Documents.Count == 0:
            raise Exception("No document is open in Word. Please open one first.")
        doc = word.ActiveDocument

        if not doc.Path:
            raise Exception(
                "The active document hasn't been saved yet. Please save it "
                "(Ctrl+S) once first — edits are applied in place to the "
                "file on disk, so it needs a real path to edit."
            )

        path = doc.FullName

        if status_cb:
            status_cb("Saving current document…")
        try:
            doc.Save()
        except Exception as e:
            raise Exception(f"Could not save the active document before editing: {e}")

        if status_cb:
            status_cb("Closing document so it can be edited in place…")
        doc.Close(SaveChanges=False)

        return path

    # ── main converter ────────────────────────────────────────────────────────
    def _docx_to_dsl(self, path, status_cb=None):
        from docx import Document

        doc = Document(path)
        lines = []

        # ── page setup ────────────────────────────────────────────────────────
        lines.append(self._extract_page_directive(doc))
        lines.append("")

        body = doc.element.body

        items = []
        body_idx = 0
        for child in body:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "p":
                from docx.text.paragraph import Paragraph

                para = Paragraph(child, doc)
                items.append(("para", para, body_idx))
                body_idx += 1
            elif tag == "tbl":
                from docx.table import Table

                tbl = Table(child, doc)
                items.append(("table", tbl, body_idx))
                body_idx += 1
            elif tag == "sectPr":
                pass

        in_list = False
        list_type = None  # "ul" or "ol"
        list_indent = 1
        pending_list_lines = []

        def flush_list_block():
            nonlocal in_list, list_type, pending_list_lines
            if in_list and pending_list_lines:
                lines.append(f"{list_type} indent={list_indent}")
                lines.extend(pending_list_lines)
                lines.append("")
            in_list = False
            list_type = None
            pending_list_lines = []

        total = len(items)
        for idx, (kind, obj, body_index) in enumerate(items):
            if status_cb and idx % 20 == 0:
                status_cb(f"Extracting element {idx+1}/{total}…")

            if kind == "table":
                flush_list_block()
                tbl_lines = self._extract_table(obj)
                if tbl_lines:
                    tbl_lines[0] = f"{tbl_lines[0]}  // id={body_index}"
                lines.extend(tbl_lines)
                lines.append("")
                continue

            para = obj

            # Page break?
            if self._is_page_break(para):
                flush_list_block()
                lines.append(f"---  // id={body_index}")
                lines.append("")
                continue

            # Horizontal rule?
            if self._has_bottom_border(para):
                flush_list_block()
                hr_line = self._extract_hr(para)
                lines.append(f"{hr_line}  // id={body_index}")
                lines.append("")
                continue

            # Blank paragraph?
            full_text = para.text.strip()

            if not full_text and not self._para_has_drawing(para):
                flush_list_block()
                lines.append(f"br  // id={body_index}")
                continue

            # Heading?
            heading_level = self._heading_level(para)
            if heading_level:
                flush_list_block()
                h_line = self._extract_heading(para, heading_level)
                lines.append(f"{h_line}  // id={body_index}")
                lines.append("")
                continue

            # List item?
            list_info = self._list_info(para)
            if list_info:
                ltype, lindent = list_info
                if not in_list or list_type != ltype or list_indent != lindent:
                    flush_list_block()
                    in_list = True
                    list_type = ltype
                    list_indent = lindent
                item_line = self._extract_list_item(para)
                pending_list_lines.append(f"{item_line}  // id={body_index}")
                continue

            # Inline drawing / image?
            if self._para_has_drawing(para):
                flush_list_block()
                lines.append(
                    f"{self._extract_drawing_placeholder(para)}  // id={body_index}"
                )
                lines.append("")
                continue

            # Regular paragraph
            flush_list_block()
            p_line = self._extract_paragraph(para)
            lines.append(f"{p_line}  // id={body_index}")
            lines.append("")

        flush_list_block()

        # Remove trailing blank lines
        while lines and lines[-1] == "":
            lines.pop()

        return "\n".join(lines)

    # ── page directive ────────────────────────────────────────────────────────
    def _extract_page_directive(self, doc):
        section = doc.sections[0]

        w_emu = section.page_width
        h_emu = section.page_height

        w_in = w_emu / INCH_TO_EMU
        h_in = h_emu / INCH_TO_EMU

        is_landscape = w_in > h_in
        if is_landscape:
            w_in, h_in = h_in, w_in

        if abs(w_in - 8.27) < 0.15 and abs(h_in - 11.69) < 0.15:
            size = "a4"
        else:
            size = "letter"

        orientation = (
            "landscape" if section.page_width > section.page_height else "portrait"
        )

        mt = round(emu_to_px(section.top_margin))
        mr = round(emu_to_px(section.right_margin))
        mb = round(emu_to_px(section.bottom_margin))
        ml = round(emu_to_px(section.left_margin))

        return f"page size={size} orientation={orientation} margin={mt},{mr},{mb},{ml}"

    # ── helpers: paragraph classification ─────────────────────────────────────
    def _heading_level(self, para):
        name = (para.style.name or "").strip()
        m = re.match(r"^[Hh]eading\s+([1-6])$", name)
        if m:
            return int(m.group(1))
        if name.lower() == "title":
            return 1
        if name.lower() == "subtitle":
            return 2
        return None

    def _list_info(self, para):
        name = (para.style.name or "").lower()
        if "list bullet" in name or "bullet" in name:
            indent = self._list_indent_from_style(name)
            return ("ul", indent)
        if "list number" in name or "list paragraph" in name:
            if self._para_is_numbered(para):
                indent = self._list_indent_from_style(name)
                return ("ol", indent)
            indent = self._list_indent_from_style(name)
            return ("ul", indent)
        numPr = para._p.find(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr"
        )
        if numPr is not None:
            ilvl_el = numPr.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl"
            )
            ilvl = (
                int(
                    ilvl_el.get(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val",
                        "0",
                    )
                )
                if ilvl_el is not None
                else 0
            )
            if self._para_is_numbered(para):
                return ("ol", ilvl + 1)
            return ("ul", ilvl + 1)
        return None

    def _list_indent_from_style(self, style_name_lower):
        m = re.search(r"(\d+)$", style_name_lower)
        if m:
            return max(1, int(m.group(1)))
        return 1

    def _para_is_numbered(self, para):
        name = (para.style.name or "").lower()
        return "number" in name or "ordered" in name

    def _is_page_break(self, para):
        for run in para.runs:
            for br in run._r.findall(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br"
            ):
                btype = br.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type",
                    "",
                )
                if btype == "page":
                    return True
        return False

    def _has_bottom_border(self, para):
        pPr = para._p.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr"
        )
        if pPr is None:
            return False
        pBdr = pPr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr"
        )
        if pBdr is None:
            return False
        bottom = pBdr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom"
        )
        return bottom is not None

    def _para_has_drawing(self, para):
        return bool(
            para._p.findall(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing"
            )
            or para._p.findall(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict"
            )
        )

    # ── helpers: color reverse-mapping ────────────────────────────────────────
    def _hex_to_dsl_token(self, hex_color):
        if not hex_color:
            return None
        hex_upper = hex_color.upper().lstrip("#")
        reverse = {}
        for token_key, alias in DSL_COLOR_ALIASES.items():
            theme_hex = THEME_COLORS.get(alias, "").upper().lstrip("#")
            if theme_hex:
                reverse[theme_hex] = token_key
        if hex_upper in reverse:
            return reverse[hex_upper]
        return f"#{hex_upper}"

    def _rgb_color_from_run(self, run):
        try:
            color = run.font.color
            if color and color.type is not None:
                rgb = color.rgb
                if rgb:
                    return f"#{str(rgb).upper()}"
        except Exception:
            pass
        rPr = run._r.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"
        )
        if rPr is not None:
            color_el = rPr.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color"
            )
            if color_el is not None:
                val = color_el.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val",
                    "",
                )
                if val and val.lower() not in ("auto", ""):
                    return f"#{val.upper()}"
        return None

    def _para_alignment_str(self, para):
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        a = para.alignment
        if a == WD_ALIGN_PARAGRAPH.CENTER:
            return "center"
        if a == WD_ALIGN_PARAGRAPH.RIGHT:
            return "right"
        if a == WD_ALIGN_PARAGRAPH.JUSTIFY:
            return "justify"
        return "left"

    # ── extractors: individual elements ───────────────────────────────────────
    def _extract_page_break_para(self, para):
        return "pagebreak"

    def _extract_hr(self, para):
        pPr = para._p.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr"
        )
        color_token = "t2"
        weight = 1.0
        if pPr is not None:
            pBdr = pPr.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr"
            )
            if pBdr is not None:
                bottom = pBdr.find(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom"
                )
                if bottom is not None:
                    c = bottom.get(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color",
                        "",
                    )
                    if c and c.lower() not in ("auto", ""):
                        color_token = self._hex_to_dsl_token(f"#{c}")
                    sz = bottom.get(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz",
                        "8",
                    )
                    try:
                        weight = round(int(sz) / 8, 2)
                    except ValueError:
                        pass
        return f"hr color={color_token} weight={weight}"

    def _extract_heading(self, para, level):
        parts = []
        align = self._para_alignment_str(para)
        if align != "left":
            parts.append(f"align={align}")

        dominant_color = None
        dominant_size = None
        dominant_font = None
        dominant_bold = None

        for run in para.runs:
            if run.text.strip():
                c = self._rgb_color_from_run(run)
                if c and dominant_color is None:
                    dominant_color = self._hex_to_dsl_token(c)
                if dominant_size is None:
                    dominant_size = self._effective_run_size_pt(run, para)
                if run.font.name and dominant_font is None:
                    dominant_font = run.font.name
                if run.font.bold is not None and dominant_bold is None:
                    dominant_bold = run.font.bold
                break

        if dominant_color:
            parts.append(f"color={dominant_color}")

        default_size = HEADING_SIZES.get(f"h{level}", 14)
        if dominant_size and dominant_size != default_size:
            parts.append(f"size={dominant_size}")
        if dominant_font:
            parts.append(f'font="{dominant_font}"')

        rich = self._runs_to_rich(para.runs, para=para)
        prefix = f"h{level}"
        if parts:
            prefix += " " + " ".join(parts)
        return f"{prefix} | {rich}"

    def _extract_paragraph(self, para):
        parts = []
        align = self._para_alignment_str(para)
        if align != "left":
            parts.append(f"align={align}")

        fmt = para.paragraph_format
        if fmt.space_before and fmt.space_before > 0:
            parts.append(f"spacing_before={round(emu_to_pt(fmt.space_before), 1)}")
        if fmt.space_after and fmt.space_after > 0:
            parts.append(f"spacing_after={round(emu_to_pt(fmt.space_after), 1)}")

        rich = self._runs_to_rich(para.runs, para=para)
        prefix = "p"
        if parts:
            prefix += " " + " ".join(parts)
        return f"{prefix} | {rich}"

    def _extract_list_item(self, para):
        rich = self._runs_to_rich(para.runs, para=para)
        return f"  item | {rich}"

    def _extract_drawing_placeholder(self, para):
        drawings = para._p.findall(
            ".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
        )
        w_px = h_px = None
        if drawings:
            try:
                cx = int(drawings[0].get("cx", 0))
                cy = int(drawings[0].get("cy", 0))
                w_px = round(emu_to_px(cx))
                h_px = round(emu_to_px(cy))
            except Exception:
                pass

        align = self._para_alignment_str(para)
        if w_px and h_px:
            return (
                f"// [EMBEDDED IMAGE — replace with: "
                f"image url=https://… width={w_px} height={h_px} align={align}]"
            )
        return f"// [EMBEDDED IMAGE — replace with: image url=https://… align={align}]"

    def _extract_table(self, tbl):
        rows = tbl.rows
        if not rows:
            return []

        num_cols = max(len(r.cells) for r in rows)
        num_rows = len(rows)

        col_widths_emu = []
        try:
            first_row = rows[0]
            for cell in first_row.cells:
                col_widths_emu.append(cell.width or 0)
        except Exception:
            col_widths_emu = [0] * num_cols

        total_w_emu = sum(col_widths_emu) or 1
        col_pcts = [f"{round(w / total_w_emu * 100)}%" for w in col_widths_emu]

        has_header = False
        header_fill_token = None
        if num_rows > 0:
            first_cells = rows[0].cells
            if first_cells:
                try:
                    first_para = first_cells[0].paragraphs[0]
                    if first_para.runs and first_para.runs[0].font.bold:
                        has_header = True
                except Exception:
                    pass
                fill_hex = self._cell_fill_hex(first_cells[0])
                if fill_hex:
                    has_header = True
                    header_fill_token = self._hex_to_dsl_token(fill_hex)

        out = []
        table_attrs = ["width=100%"]
        if header_fill_token:
            table_attrs.append(f"header_fill={header_fill_token}")

        border_hex = self._cell_border_hex(rows[0].cells[0]) if rows else None
        if border_hex:
            border_token = self._hex_to_dsl_token(border_hex)
            table_attrs.append(f"border_color={border_token} border_weight=1")

        out.append("table " + " ".join(table_attrs))
        out.append("cols=" + ",".join(col_pcts))

        start_row = 0
        if has_header:
            header_cells = [c.text.strip() for c in rows[0].cells]
            out.append("header=" + ",".join(f'"{v}"' for v in header_cells))
            start_row = 1

        for ri in range(start_row, num_rows):
            cells = [c.text.strip() for c in rows[ri].cells]
            out.append("row=" + ",".join(f'"{v}"' for v in cells))

        return out

    def _cell_fill_hex(self, cell):
        tc = cell._tc
        tcPr = tc.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr"
        )
        if tcPr is None:
            return None
        shd = tcPr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd"
        )
        if shd is None:
            return None
        fill = shd.get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill", ""
        )
        if fill and fill.lower() not in ("auto", "ffffff", ""):
            return f"#{fill.upper()}"
        return None

    def _cell_border_hex(self, cell):
        tc = cell._tc
        tcPr = tc.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr"
        )
        if tcPr is None:
            return None
        tcBorders = tcPr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders"
        )
        if tcBorders is None:
            return None
        top = tcBorders.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top"
        )
        if top is None:
            return None
        c = top.get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color", ""
        )
        if c and c.lower() not in ("auto", ""):
            return f"#{c.upper()}"
        return None

    # ── rich text run serializer ───────────────────────────────────────────────
    def _runs_to_rich(self, runs, para=None):
        segments = []
        for run in runs:
            text = run.text
            if not text:
                continue
            seg = {"text": text}
            if run.font.bold:
                seg["bold"] = True
            if run.font.italic:
                seg["italic"] = True
            if run.font.underline:
                seg["underline"] = True

            if para is not None:
                pt = self._effective_run_size_pt(run, para)
            else:
                pt = round(emu_to_pt(run.font.size)) if run.font.size else None
            if pt and pt != 11:
                seg["size"] = pt

            if run.font.name:
                seg["font"] = run.font.name
            c = self._rgb_color_from_run(run)
            if c:
                seg["color"] = self._hex_to_dsl_token(c)
            segments.append(seg)

        if not segments:
            return '""'

        merged = [segments[0]]
        for seg in segments[1:]:
            prev = merged[-1]
            same = (
                prev.get("bold") == seg.get("bold")
                and prev.get("italic") == seg.get("italic")
                and prev.get("underline") == seg.get("underline")
                and prev.get("size") == seg.get("size")
                and prev.get("font") == seg.get("font")
                and prev.get("color") == seg.get("color")
            )
            if same:
                prev["text"] += seg["text"]
            else:
                merged.append(seg)

        parts = []
        for seg in merged:
            text = seg["text"].replace('"', '\\"')
            attrs = []
            if seg.get("bold"):
                attrs.append("bold=true")
            if seg.get("italic"):
                attrs.append("italic=true")
            if seg.get("underline"):
                attrs.append("underline=true")
            if seg.get("size"):
                attrs.append(f"size={seg['size']}")
            if seg.get("font"):
                attrs.append(f'font="{seg["font"]}"')
            if seg.get("color"):
                attrs.append(f"color={seg['color']}")
            if attrs:
                parts.append(f'"{text}" {" ".join(attrs)}')
            else:
                parts.append(f'"{text}"')

        return " + ".join(parts)

    def _effective_run_size_pt(self, run, para):
        from docx.oxml.ns import qn

        if run.font.size:
            return round(emu_to_pt(run.font.size))

        style = para.style
        while style is not None:
            rPr = style.element.find(qn("w:rPr"))
            if rPr is not None:
                sz = rPr.find(qn("w:sz"))
                if sz is not None:
                    val = sz.get(qn("w:val"))
                    if val:
                        try:
                            return round(int(val) / 2)
                        except ValueError:
                            pass
            try:
                style = style.base_style
            except Exception:
                break

        try:
            styles_part = para.part.styles
            doc_defaults = styles_part._element.find(qn("w:docDefaults"))
            if doc_defaults is not None:
                rPrDefault = doc_defaults.find(
                    f"{qn('w:rPrDefault')}/{qn('w:rPr')}"
                )
                if rPrDefault is not None:
                    sz = rPrDefault.find(qn("w:sz"))
                    if sz is not None:
                        val = sz.get(qn("w:val"))
                        if val:
                            try:
                                return round(int(val) / 2)
                            except ValueError:
                                pass
        except Exception:
            pass

        return None
