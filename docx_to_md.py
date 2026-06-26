#!/usr/bin/env python3
"""Convert a Word (.docx) document to Markdown (.md).

Preserves bold, italic, strikethrough, inline code, headings, bulleted and
numbered lists (including nesting), hyperlinks, images, and tables. Tables with
merged or nested cells fall back to raw HTML (which Markdown permits).

Text in a monospace font is rendered as inline `code`, but only as a heuristic:
for a document that is *predominantly* monospace (a log or terminal dump), that
signal is meaningless, so it is disabled automatically and bold/italic render
normally. Use --code-detection to force the behaviour either way.

Footnotes, comments, and tracked-change marks are dropped; the final accepted
text is kept (insertions are accepted, deletions removed).

Usage:
    python docx_to_md.py input.docx -o output.md
    python docx_to_md.py input.docx --html-tables --verbose
    python docx_to_md.py input.docx --code-detection off
    python docx_to_md.py input.docx --engine mammoth

See README.md for the full feature list and known limitations.
"""

from __future__ import annotations

import argparse
import os
import re
import sys

from docx import Document
from docx.document import Document as _Document
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph


# Fonts that should be rendered as inline `code`.
MONOSPACE_FONTS = {
    "courier", "courier new", "consolas", "lucida console", "monospace",
    "cascadia code", "cascadia mono", "dejavu sans mono", "menlo", "monaco",
    "source code pro", "sf mono", "roboto mono", "fira code", "fira mono",
    "liberation mono", "inconsolata", "andale mono", "ibm plex mono",
}

# Markdown special characters that must be escaped when they appear literally.
_ESCAPE_RE = re.compile(r"([\\`*_\[\]])")


def escape_md(text: str) -> str:
    """Escape Markdown-significant characters in plain text."""
    return _ESCAPE_RE.sub(r"\\\1", text)


def escape_html(text: str) -> str:
    """Escape characters significant in HTML element content."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def escape_attr(text: str) -> str:
    """Escape a string for use inside an HTML attribute value."""
    return escape_html(text).replace('"', "&quot;")


def iter_block_items(parent):
    """Yield Paragraph and Table objects from *parent* in document order.

    *parent* may be a Document or a table cell. Paragraphs and tables interleave
    in the body, and order matters, so we walk the XML children directly.
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("unsupported parent type: %r" % type(parent))

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _emphasize(text: str, bold: bool, italic: bool, strike: bool) -> str:
    """Wrap *text* in Markdown emphasis markers, keeping whitespace outside.

    Markdown does not render emphasis when a marker hugs a space (``** x **``),
    so leading/trailing whitespace is moved outside the markers.
    """
    if not text.strip():
        return text
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()):]
    core = text.strip()
    if strike:
        core = "~~%s~~" % core
    if bold and italic:
        core = "***%s***" % core
    elif bold:
        core = "**%s**" % core
    elif italic:
        core = "*%s*" % core
    return lead + core + trail


def _code_span(text: str) -> str:
    """Wrap *text* in a Markdown code span, handling embedded backticks."""
    if "`" in text:
        return "`` %s ``" % text
    return "`%s`" % text


class Converter:
    """Walks a python-docx Document and emits Markdown."""

    def __init__(self, docx_path, output_path, images_dir="images",
                 force_html_tables=False, code_detection="auto", verbose=False):
        self.document = Document(docx_path)
        self.output_dir = os.path.dirname(os.path.abspath(output_path))
        if os.path.isabs(images_dir):
            self.images_dir = images_dir
        else:
            self.images_dir = os.path.join(self.output_dir, images_dir)
        self.force_html_tables = force_html_tables
        self.code_detection = code_detection
        self.verbose = verbose
        self.warnings = []
        self._image_cache = {}   # rId -> relative path written into the markdown
        self._used_names = {}     # filename -> rId (collision tracking)
        self._num_fmt = {}        # numId(str) -> {ilvl(int): numFmt(str)}
        self._build_numbering_map()
        self._code_enabled = self._resolve_code_detection()

    # ----- warnings ---------------------------------------------------------

    def warn(self, message):
        self.warnings.append(message)
        if self.verbose:
            print("warning: %s" % message, file=sys.stderr)

    # ----- numbering --------------------------------------------------------

    def _build_numbering_map(self):
        """Resolve numId -> {level: numFmt} so lists can be ordered vs bullet."""
        try:
            numbering = self.document.part.numbering_part.element
        except (NotImplementedError, KeyError, AttributeError):
            return

        abstract = {}  # abstractNumId -> {ilvl: numFmt}
        for an in numbering.findall(qn("w:abstractNum")):
            aid = an.get(qn("w:abstractNumId"))
            levels = {}
            for lvl in an.findall(qn("w:lvl")):
                ilvl = lvl.get(qn("w:ilvl"))
                fmt_el = lvl.find(qn("w:numFmt"))
                fmt = fmt_el.get(qn("w:val")) if fmt_el is not None else None
                try:
                    levels[int(ilvl)] = fmt
                except (TypeError, ValueError):
                    pass
            abstract[aid] = levels

        for num in numbering.findall(qn("w:num")):
            num_id = num.get(qn("w:numId"))
            ref = num.find(qn("w:abstractNumId"))
            aid = ref.get(qn("w:val")) if ref is not None else None
            self._num_fmt[num_id] = abstract.get(aid, {})

    # ----- top-level conversion --------------------------------------------

    def convert(self):
        blocks = []
        list_lines = []

        def flush_list():
            if list_lines:
                blocks.append("\n".join(list_lines))
                list_lines.clear()

        for block in iter_block_items(self.document):
            if isinstance(block, Paragraph):
                info = self._list_info(block)
                if info is not None:
                    line = self._render_list_item(block, info)
                    if line is not None:
                        list_lines.append(line)
                    continue
                flush_list()
                rendered = self._render_paragraph(block)
                if rendered:
                    blocks.append(rendered)
            else:  # Table
                flush_list()
                rendered = self._render_table(block)
                if rendered:
                    blocks.append(rendered)

        flush_list()

        md = "\n\n".join(blocks).strip()
        # Collapse runs of 3+ newlines down to a single blank line.
        md = re.sub(r"\n{3,}", "\n\n", md)
        return md + "\n" if md else ""

    # ----- paragraphs -------------------------------------------------------

    def _style_name(self, paragraph):
        try:
            return paragraph.style.name or ""
        except Exception:
            return ""

    def _render_paragraph(self, paragraph):
        content = self._render_children(paragraph._p).strip()
        style = self._style_name(paragraph)

        # Horizontal rule: an empty paragraph carrying a border.
        if not content:
            pPr = paragraph._p.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:pBdr")) is not None:
                return "---"
            return ""

        heading = re.match(r"Heading\s+(\d+)", style)
        if heading:
            level = min(int(heading.group(1)), 6)
            return "#" * level + " " + content
        if style == "Title":
            return "# " + content
        if style == "Subtitle":
            return "## " + content
        if style in ("Quote", "Intense Quote"):
            return "\n".join("> " + line for line in content.splitlines())

        # Soft line breaks inside a paragraph become hard breaks in Markdown.
        return content.replace("\n", "  \n")

    # ----- lists ------------------------------------------------------------

    def _list_info(self, paragraph):
        """Return (level, kind) where kind is 'bullet' or 'ordered', else None."""
        pPr = paragraph._p.find(qn("w:pPr"))
        if pPr is None:
            return self._list_info_from_style(paragraph)
        numPr = pPr.find(qn("w:numPr"))
        if numPr is None:
            return self._list_info_from_style(paragraph)

        ilvl_el = numPr.find(qn("w:ilvl"))
        num_id_el = numPr.find(qn("w:numId"))
        try:
            ilvl = int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0
        except (TypeError, ValueError):
            ilvl = 0
        num_id = num_id_el.get(qn("w:val")) if num_id_el is not None else None

        # numId 0 conventionally means "no list".
        if num_id in (None, "0"):
            return self._list_info_from_style(paragraph)

        fmt = self._num_fmt.get(num_id, {}).get(ilvl)
        kind = "bullet" if (fmt is None or fmt == "bullet") else "ordered"
        return (ilvl, kind)

    def _list_info_from_style(self, paragraph):
        name = self._style_name(paragraph).lower()
        # Built-in list styles encode the nesting level in the name itself,
        # e.g. "List Bullet" (level 0), "List Bullet 2" (level 1), etc.
        match = re.search(r"(\d+)\s*$", name)
        level = max(int(match.group(1)) - 1, 0) if match else 0
        if "list bullet" in name:
            return (level, "bullet")
        if "list number" in name:
            return (level, "ordered")
        return None

    def _render_list_item(self, paragraph, info):
        level, kind = info
        content = self._render_children(paragraph._p).strip()
        if not content:
            return None
        indent = "  " * level
        marker = "1. " if kind == "ordered" else "- "
        # Keep multi-line cell content on one logical line.
        content = content.replace("\n", " ")
        return indent + marker + content

    # ----- inline content ---------------------------------------------------

    def _render_children(self, element, html=False):
        """Render the inline children of *element* (a paragraph or wrapper)."""
        out = []
        for child in element:
            tag = child.tag
            if tag == qn("w:r"):
                out.append(self._render_run(child, html))
            elif tag == qn("w:hyperlink"):
                out.append(self._render_hyperlink(child, html))
            elif tag == qn("w:ins"):
                # Tracked insertion: accept it (render the inner content).
                out.append(self._render_children(child, html))
            elif tag == qn("w:del"):
                # Tracked deletion: accept it (drop the content).
                continue
            elif tag == qn("w:smartTag"):
                out.append(self._render_children(child, html))
            elif tag == qn("w:fldSimple"):
                out.append(self._render_fldsimple(child, html))
            # Everything else (pPr, bookmarks, comment refs, etc.) is ignored.
        return "".join(out)

    def _bool_prop(self, rPr, tag):
        if rPr is None:
            return False
        el = rPr.find(qn(tag))
        if el is None:
            return False
        val = el.get(qn("w:val"))
        if val is None:
            return True
        return val.lower() not in ("0", "false", "off", "none")

    # ----- code-span detection ---------------------------------------------

    @staticmethod
    def _font_of(rPr):
        """Return the run's directly-set font name, or None if inherited."""
        if rPr is None:
            return None
        rfonts = rPr.find(qn("w:rFonts"))
        if rfonts is None:
            return None
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            font = rfonts.get(qn(attr))
            if font:
                return font
        return None

    def _default_font(self):
        """Resolve the document's default body font (docDefaults, then Normal)."""
        name = None
        try:
            styles_el = self.document.styles.element
        except Exception:
            return None
        dd = styles_el.find(qn("w:docDefaults"))
        if dd is not None:
            rprd = dd.find(qn("w:rPrDefault"))
            rpr = rprd.find(qn("w:rPr")) if rprd is not None else None
            name = self._font_of(rpr) or name
        try:
            normal = self.document.styles["Normal"]
            name = self._font_of(normal.element.find(qn("w:rPr"))) or name
        except Exception:
            pass
        return name

    def _resolve_code_detection(self):
        """Decide whether monospace runs should become inline code.

        ``on``/``off`` are explicit. ``auto`` keeps the heuristic on for normal
        prose but turns it off when the document is *predominantly* monospace
        (e.g. a terminal/log dump), where a monospace font signals nothing.
        """
        if self.code_detection == "on":
            return True
        if self.code_detection == "off":
            return False

        default = self._default_font()
        default_mono = bool(default) and default.lower() in MONOSPACE_FONTS
        mono = total = 0
        for r in self.document.element.body.iter(qn("w:r")):
            text = self._run_text(r)
            if not text:
                continue
            total += len(text)
            font = self._font_of(r.find(qn("w:rPr")))
            is_mono = default_mono if font is None else font.lower() in MONOSPACE_FONTS
            if is_mono:
                mono += len(text)

        dominant = default_mono if total == 0 else mono >= 0.5 * total
        if dominant:
            self.warn("monospace is the body font; inline-code detection "
                      "disabled (use --code-detection on to force it)")
        return not dominant

    def _is_code(self, rPr):
        if not self._code_enabled:
            return False
        font = self._font_of(rPr)
        return font is not None and font.lower() in MONOSPACE_FONTS

    def _run_text(self, run_el):
        parts = []
        for el in run_el:
            tag = el.tag
            if tag == qn("w:t"):
                parts.append(el.text or "")
            elif tag == qn("w:tab"):
                parts.append("\t")
            elif tag in (qn("w:br"), qn("w:cr")):
                parts.append("\n")
            elif tag == qn("w:noBreakHyphen"):
                parts.append("-")
            # w:delText is intentionally skipped (tracked deletion).
        return "".join(parts)

    def _render_run(self, run_el, html=False):
        out = ""
        text = self._run_text(run_el)
        if text:
            rPr = run_el.find(qn("w:rPr"))
            bold = self._bool_prop(rPr, "w:b")
            italic = self._bool_prop(rPr, "w:i")
            strike = self._bool_prop(rPr, "w:strike")
            code = self._is_code(rPr)
            if html:
                out += self._format_html(text, bold, italic, strike, code)
            elif code:
                out += _code_span(text)
            else:
                out += _emphasize(escape_md(text), bold, italic, strike)

        # Embedded images (DrawingML and legacy VML).
        for blip in run_el.findall(".//" + qn("a:blip")):
            rid = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
            out += self._image_md(rid, run_el, html)
        try:
            for imagedata in run_el.findall(".//" + qn("v:imagedata")):
                out += self._image_md(imagedata.get(qn("r:id")), run_el, html)
        except KeyError:
            pass  # 'v' namespace unavailable in this build
        return out

    def _format_html(self, text, bold, italic, strike, code):
        if code:
            return "<code>%s</code>" % escape_html(text)
        if not text.strip():
            return escape_html(text)
        lead = text[: len(text) - len(text.lstrip())]
        trail = text[len(text.rstrip()):]
        core = escape_html(text.strip())
        if strike:
            core = "<del>%s</del>" % core
        if italic:
            core = "<em>%s</em>" % core
        if bold:
            core = "<strong>%s</strong>" % core
        return lead + core + trail

    def _rel_target(self, rid):
        try:
            return self.document.part.rels[rid].target_ref
        except KeyError:
            return None

    def _render_hyperlink(self, link_el, html=False):
        inner = self._render_children(link_el, html).strip()
        rid = link_el.get(qn("r:id"))
        anchor = link_el.get(qn("w:anchor"))
        url = self._rel_target(rid) if rid else None
        if url is None and anchor:
            url = "#" + anchor
        if not inner:
            inner = url or ""
        if not url:
            return inner
        if html:
            return '<a href="%s">%s</a>' % (escape_attr(url), inner)
        return "[%s](%s)" % (inner, url)

    def _render_fldsimple(self, el, html=False):
        instr = el.get(qn("w:instr")) or ""
        inner = self._render_children(el, html).strip()
        match = re.search(r'HYPERLINK\s+"([^"]+)"', instr)
        if match:
            url = match.group(1)
            inner = inner or url
            if html:
                return '<a href="%s">%s</a>' % (escape_attr(url), inner)
            return "[%s](%s)" % (inner, url)
        return inner

    # ----- images -----------------------------------------------------------

    def _image_alt(self, container_el):
        docpr = container_el.find(".//" + qn("wp:docPr"))
        if docpr is not None:
            return docpr.get("descr") or docpr.get("name") or "image"
        return "image"

    def _image_md(self, rid, container_el, html=False):
        if not rid:
            return ""
        alt = self._image_alt(container_el)
        rel = self._image_cache.get(rid)
        if rel is None:
            try:
                image_part = self.document.part.related_parts[rid]
            except KeyError:
                self.warn("image relationship %s not found; skipped" % rid)
                return ""
            name = os.path.basename(str(image_part.partname)) or "image"
            # Avoid overwriting a different image that shares a basename.
            if name in self._used_names and self._used_names[name] != rid:
                root, ext = os.path.splitext(name)
                name = "%s_%s%s" % (root, len(self._used_names), ext)
            self._used_names[name] = rid

            os.makedirs(self.images_dir, exist_ok=True)
            dest = os.path.join(self.images_dir, name)
            with open(dest, "wb") as fh:
                fh.write(image_part.blob)
            rel = os.path.relpath(dest, self.output_dir).replace("\\", "/")
            self._image_cache[rid] = rel
            if self.verbose:
                print("extracted image: %s" % rel, file=sys.stderr)

        if html:
            return '<img src="%s" alt="%s">' % (escape_attr(rel), escape_attr(alt))
        return "![%s](%s)" % (alt.replace("]", ""), rel)

    # ----- tables -----------------------------------------------------------

    def _table_needs_html(self, tbl_el):
        if self.force_html_tables:
            return True
        for tc in tbl_el.findall(".//" + qn("w:tc")):
            tcpr = tc.find(qn("w:tcPr"))
            if tcpr is not None:
                if tcpr.find(qn("w:gridSpan")) is not None:
                    return True
                if tcpr.find(qn("w:vMerge")) is not None:
                    return True
            if tc.find(qn("w:tbl")) is not None:
                return True
        return False

    def _render_table(self, table):
        tbl_el = table._tbl
        if self._table_needs_html(tbl_el):
            reason = "forced" if self.force_html_tables else "merged or nested cells"
            self.warn("table rendered as HTML (%s)" % reason)
            note = "<!-- note: table rendered as HTML due to %s -->" % reason
            return note + "\n" + self._render_html_table_el(tbl_el)
        return self._render_pipe_table(table)

    def _cell_text(self, cell):
        parts = []
        for para in cell.paragraphs:
            text = self._render_children(para._p).strip()
            if text:
                parts.append(text)
        joined = "<br>".join(parts)
        return joined.replace("\n", " ").replace("|", "\\|")

    def _render_pipe_table(self, table):
        rows = table.rows
        if not rows:
            return ""
        header = [self._cell_text(c) for c in rows[0].cells]
        n = len(header)
        lines = [
            "| " + " | ".join(header) + " |",
            "| " + " | ".join(["---"] * n) + " |",
        ]
        for row in rows[1:]:
            cells = [self._cell_text(c) for c in row.cells]
            while len(cells) < n:
                cells.append("")
            lines.append("| " + " | ".join(cells[:n]) + " |")
        return "\n".join(lines)

    def _render_tc_html(self, tc):
        parts = []
        for child in tc:
            if child.tag == qn("w:p"):
                text = self._render_children(child, html=True).strip()
                if text:
                    parts.append(text)
            elif child.tag == qn("w:tbl"):
                parts.append(self._render_html_table_el(child))
        return "<br>".join(parts)

    def _render_html_table_el(self, tbl_el):
        rows = tbl_el.findall(qn("w:tr"))
        grid = []            # rendered rows; each is a list of cell dicts
        active_vmerge = {}    # grid column index -> cell dict spanning down

        for tr in rows:
            row_cells = []
            col = 0
            for tc in tr.findall(qn("w:tc")):
                tcpr = tc.find(qn("w:tcPr"))
                gridspan = 1
                vmerge = None
                if tcpr is not None:
                    gs = tcpr.find(qn("w:gridSpan"))
                    if gs is not None:
                        try:
                            gridspan = int(gs.get(qn("w:val")))
                        except (TypeError, ValueError):
                            gridspan = 1
                    vm = tcpr.find(qn("w:vMerge"))
                    if vm is not None:
                        vmerge = (vm.get(qn("w:val")) or "continue").lower()

                if vmerge == "continue":
                    ref = active_vmerge.get(col)
                    if ref is not None:
                        ref["rowspan"] += 1
                    col += gridspan
                    continue

                cell = {
                    "content": self._render_tc_html(tc),
                    "colspan": gridspan,
                    "rowspan": 1,
                }
                row_cells.append(cell)
                if vmerge == "restart":
                    active_vmerge[col] = cell
                else:
                    active_vmerge.pop(col, None)
                col += gridspan
            grid.append(row_cells)

        out = ["<table>"]
        for i, row_cells in enumerate(grid):
            out.append("  <tr>")
            tag = "th" if i == 0 else "td"
            for cell in row_cells:
                attrs = ""
                if cell["colspan"] > 1:
                    attrs += ' colspan="%d"' % cell["colspan"]
                if cell["rowspan"] > 1:
                    attrs += ' rowspan="%d"' % cell["rowspan"]
                out.append("    <%s%s>%s</%s>" % (tag, attrs, cell["content"], tag))
            out.append("  </tr>")
        out.append("</table>")
        return "\n".join(out)


def docx_to_md(path, output_path, images_dir="images",
               force_html_tables=False, code_detection="auto", verbose=False):
    """Convert *path* to Markdown using the custom python-docx engine.

    Returns (markdown_text, warnings).
    """
    converter = Converter(path, output_path, images_dir=images_dir,
                          force_html_tables=force_html_tables,
                          code_detection=code_detection, verbose=verbose)
    md = converter.convert()
    return md, converter.warnings


def docx_to_md_baseline(path):
    """Quick baseline conversion using mammoth (used for comparison/fallback)."""
    import mammoth
    with open(path, "rb") as fh:
        result = mammoth.convert_to_markdown(fh)
    return result.value


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Convert a Word (.docx) document to Markdown (.md).")
    parser.add_argument("input", help="path to the input .docx file")
    parser.add_argument("-o", "--output",
                        help="output .md path (default: input name with .md)")
    parser.add_argument("--images-dir", default="images",
                        help="directory for extracted images, relative to the "
                             "output file (default: images)")
    parser.add_argument("--html-tables", action="store_true",
                        help="render every table as HTML, not just merged ones")
    parser.add_argument("--code-detection", choices=("auto", "on", "off"),
                        default="auto",
                        help="treat monospace-font text as inline code: 'auto' "
                             "(default) disables it for all-monospace documents, "
                             "'on' always, 'off' never")
    parser.add_argument("--engine", choices=("custom", "mammoth"),
                        default="custom",
                        help="conversion engine (default: custom)")
    parser.add_argument("--verbose", action="store_true",
                        help="print warnings and progress to stderr")
    args = parser.parse_args(argv)

    if not os.path.isfile(args.input):
        parser.error("input file not found: %s" % args.input)
    if args.input.lower().endswith(".doc") and not args.input.lower().endswith(".docx"):
        parser.error("legacy .doc files are not supported; convert to .docx first")

    output_path = args.output or (os.path.splitext(args.input)[0] + ".md")

    if args.engine == "mammoth":
        md = docx_to_md_baseline(args.input)
        warnings = []
    else:
        md, warnings = docx_to_md(
            args.input, output_path,
            images_dir=args.images_dir,
            force_html_tables=args.html_tables,
            code_detection=args.code_detection,
            verbose=args.verbose,
        )

    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write(md)

    print("wrote %s" % output_path)
    if warnings and not args.verbose:
        print("%d warning(s); re-run with --verbose for details." % len(warnings),
              file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
