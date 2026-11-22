#!/usr/bin/env python3
"""Build an editable company report from a local JSON layout manifest."""

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlsplit

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Pt, RGBColor, Twips


ASSET = Path(__file__).resolve().parent.parent / "assets" / "report-shell.docx"

red = "BC131B"
black = "171717"
grey = "565453"
role = {
    "title": (24, 30, 0, 8, red, True),
    "subtitle": (17.5, 24, 0, 6, red, True),
    "assignment": (9.5, 14, 0, 11, grey, False),
    "section": (12.2, 17, 6, 5, red, True),
    "paragraph": (10.75, 15.7, 0, 7, black, False),
    "note": (8.1, 11.5, 2, 5, grey, False),
    "sources": (8.1, 11.5, 3, 4, grey, False),
    "band": (11.5, 18, 0, 8, "FFFFFF", True),
}


def xml(tag, **attributes):
    element = OxmlElement("w:" + tag)
    for key, value in attributes.items():
        element.set(qn("w:" + key), str(value))
    return element


def font(run, size, color, bold=False):
    run.font.name = "Tahoma"
    r = run._element.get_or_add_rPr()
    r.rFonts.set(qn("w:eastAsia"), "楷体_GB2312")
    r.rFonts.set(qn("w:cs"), "Tahoma")
    r.append(xml("lang", val="en-US", eastAsia="zh-CN"))
    run.font.color.rgb = RGBColor.from_string(color)
    run.bold = bold
    run.font.size = Pt(size)


def paragraph_format(p, leading, before, after):
    f = p.paragraph_format
    f.space_before = Pt(before)
    f.space_after = Pt(after)
    f.line_spacing = Pt(leading)
    f.keep_with_next = False
    f.keep_together = True
    f.widow_control = False
    f.left_indent = Pt(0)
    f.right_indent = Pt(0)
    f.first_line_indent = Pt(0)


def p_border(p, key="bottom", stroke=red, weight=8, space=1):
    x = xml("pBdr")
    x.append(xml(key, val="single", sz=weight, space=space, color=stroke))
    p._p.get_or_add_pPr().append(x)


def footer(p, width, company, as_of):
    paragraph_format(p, 12, 0, 0)
    p.style = "Normal"
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.tab_stops.add_tab_stop(Pt(width), WD_TAB_ALIGNMENT.RIGHT)
    p_border(p, "top", weight=7, space=3)
    font(p.add_run("P."), 9, red, True)
    f = xml("fldSimple", instr=" PAGE ")
    r = xml("r")
    r.append(xml("t"))
    r[0].text = "1"
    f.append(r)
    p._p.append(f)
    font(p.add_run("\t" + "  |  ".join(filter(None, (company, as_of)))), 8.5, grey)


def set_headers(section, width, metadata):
    company = metadata["company"]
    as_of = metadata.get("as_of", "")
    series = metadata.get("series", "公司研究")
    kind = metadata.get("report_kind", "深度研究")
    author = metadata.get("author", "")
    heading = metadata.get("masthead", series)
    logo = metadata.get("logo")

    default_header = section.header.paragraphs[0]
    paragraph_format(default_header, 14, 0, 0)
    default_header.style = "Normal"
    default_header.alignment = WD_ALIGN_PARAGRAPH.LEFT
    default_header.paragraph_format.tab_stops.add_tab_stop(
        Pt(width), WD_TAB_ALIGNMENT.RIGHT
    )
    font(default_header.add_run(series), 10, black, True)
    font(
        default_header.add_run("\t" + "  |  ".join(filter(None, (author, as_of)))),
        9,
        grey,
    )
    p_border(default_header, weight=8, space=4)
    footer(section.footer.paragraphs[0], width, company, as_of)

    section.different_first_page_header_footer = True
    section.top_margin = Twips(1301)
    masthead = section.first_page_header.paragraphs[0]
    paragraph_format(masthead, 25, 0, 0)
    masthead.style = "Normal"
    masthead.alignment = WD_ALIGN_PARAGRAPH.LEFT
    masthead.paragraph_format.tab_stops.add_tab_stop(Pt(width), WD_TAB_ALIGNMENT.RIGHT)
    if logo:
        masthead.paragraph_format.line_spacing = 1
        masthead.add_run().add_picture(logo, height=Pt(29))
        heading = "  " + heading
    font(masthead.add_run(heading), 13.5, black, True)
    right = "  ·  ".join(dict.fromkeys(filter(None, (series, kind))))
    font(masthead.add_run("\t" + right), 9.5, red)
    p_border(masthead, weight=12, space=4)
    footer(section.first_page_footer.paragraphs[0], width, company, as_of)


def link(par, name, url, size):
    h = xml("hyperlink")
    h.set(qn("r:id"), par.part.relate_to(url, RT.HYPERLINK, is_external=True))
    r = par.add_run(name)
    font(r, size, red)
    r.underline = True
    h.append(r._r)
    par._p.append(h)


def empty(par):
    paragraph_format(par, 1, 0, 0)
    font(par.add_run("\u00a0"), 1, black)


def append(container, block, available):
    typ = block["type"]
    if typ == "columns":
        gap = block.get("gap", 12)
        cols = [
            (available - gap) * block["fractions"][0],
            gap,
            (available - gap) * block["fractions"][1],
        ]
        frame = make_table(container, 1, cols)
        r = frame.rows[0]
        for idx, b in zip((0, 2), block["blocks"]):
            cell = r.cells[idx]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            run_blocks(cell, b, cols[idx])
        empty(r.cells[1].paragraphs[0])
        return
    if typ in ("table", "kv"):
        data = [block["headers"], *block["rows"]] if typ == "table" else block["rows"]
        frac = block.get("fractions", [0.64, 0.36])
        frame = make_table(container, len(data), [available * f for f in frac])
        head = typ == "table"
        for i, vals in enumerate(data):
            for j, (c, t) in enumerate(zip(frame.rows[i].cells, vals)):
                set_padding(c, 40, 64)
                vertical_line(
                    c,
                    "bottom",
                    red if head and (i == 0 or i == len(data) - 1) else "E8DDDD",
                    6 if head and (i == 0 or i == len(data) - 1) else 3,
                )
                if head and i == 0:
                    vertical_line(c, "top", red, 7)
                pp = c.paragraphs[0]
                paragraph_format(pp, 11.5 if available < 350 else 12.5, 0, 0)
                text_col = j == 0 or (i and j in block.get("text_cols", ()))
                pp.alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT if text_col else WD_ALIGN_PARAGRAPH.RIGHT
                )
                sz = 8.4 if available < 350 else 9.1
                if head and i == 0 and len(frac) >= 6:
                    sz = 8.0
                font(
                    pp.add_run(str(t)),
                    sz,
                    red if head and i == 0 else black,
                    head and i == 0,
                )
        return
    if typ == "image":
        pp = container.add_paragraph()
        pf = pp.paragraph_format
        pf.space_before, pf.space_after = Pt(3), Pt(3)
        pf.line_spacing = 1
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp.add_run().add_picture(block["path"], width=Pt(available - 3))
        return
    pp = container.add_paragraph()
    sz, lead, before, after, col, bold = role[typ]
    if typ == "section" and available < 350:
        sz, lead, before, after = (10, 14, 6, 4)
        p_border(pp, "top", weight=6, space=2)
    paragraph_format(pp, lead, before, after)
    if typ == "band":
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp._p.get_or_add_pPr().append(xml("shd", fill=red, val="clear"))
    if typ in ("title", "subtitle", "section"):
        pp.paragraph_format.keep_with_next = True
    if typ == "sources":
        font(pp.add_run("资料来源  "), sz, grey)
        for i, (lab, url) in enumerate(block["items"]):
            if i:
                font(pp.add_run("  /  "), sz, grey)
            link(pp, lab, url, sz)
        return
    if block.get("lead"):
        font(pp.add_run(block["lead"] + " "), sz, red, True)
    if block.get("text"):
        font(pp.add_run(block["text"]), sz, col, bold)


def vertical_line(cell, key, color, weight):
    pr = cell._tc.get_or_add_tcPr()
    borders = pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = xml("tcBorders")
        pr.append(borders)
    borders.append(xml(key, val="single", sz=weight, color=color))


def set_padding(cell, v, h):
    pr = cell._tc.get_or_add_tcPr()
    m = pr.find(qn("w:tcMar"))
    if m is not None:
        pr.remove(m)
    m = xml("tcMar")
    for n, z in [("top", v), ("bottom", v), ("left", h), ("right", h)]:
        m.append(xml(n, w=z, type="dxa"))
    pr.append(m)


def make_table(container, rows, widths):
    table = container.add_table(rows, len(widths))
    if hasattr(container, "_tc"):
        empty(container.paragraphs[-1])
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tblw = table._tbl.tblPr.find(qn("w:tblW"))
    tblw.set(qn("w:w"), str(round(sum(widths) * 20)))
    tblw.set(qn("w:type"), "dxa")
    table._tbl.tblPr.append(xml("tblBorders"))
    for column, w in zip(table.columns, widths):
        column.width = Pt(w)
    for row in table.rows:
        row._tr.get_or_add_trPr().append(xml("cantSplit"))
        for cell, w in zip(row.cells, widths):
            cell.width = Pt(w)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_padding(cell, 0, 0)
    return table


def run_blocks(container, blocks, available):
    first = container.paragraphs[0] if hasattr(container, "_tc") else None
    for b in blocks:
        append(container, b, available)
    if first is not None:
        first._element.getparent().remove(first._element)
        terminal = container._tc[-1]
        if terminal.tag != qn("w:p"):
            empty(container.add_paragraph())
        elif not container.paragraphs[-1].text and not terminal.xpath(".//w:drawing"):
            empty(container.paragraphs[-1])


def local_file(base, supplied):
    file = Path(supplied).expanduser()
    file = (base / file).resolve() if not file.is_absolute() else file
    if not file.is_file():
        raise ValueError(f"Missing local image: {file}")
    return str(file)


def normalize_widths(block, count, fallback=None):
    values = block.get("fractions", fallback)
    if len(values) != count or any(
        not isinstance(x, (float, int)) or x <= 0 for x in values
    ):
        raise ValueError(f"{block['type']} requires {count} positive column widths")
    total = sum(values)
    block["fractions"] = [v / total for v in values]


def prepare_blocks(blocks, root):
    if not isinstance(blocks, list) or not blocks:
        raise ValueError("A page or column must contain a nonempty block list")
    for block in blocks:
        kind = block["type"]
        if kind == "columns":
            parts = block["blocks"]
            if len(parts) != 2:
                raise ValueError("A columns block must contain two block lists")
            normalize_widths(block, 2, [0.5, 0.5])
            for column in parts:
                prepare_blocks(column, root)
        elif kind == "table":
            count = len(block["headers"])
            if count < 2 or any(len(row) != count for row in block["rows"]):
                raise ValueError(
                    "A table needs consistent rows and at least two headers"
                )
            normalize_widths(block, count, [1] * count)
        elif kind == "kv":
            if not block["rows"] or any(len(row) != 2 for row in block["rows"]):
                raise ValueError("A kv block requires label/value pairs")
            normalize_widths(block, 2, [0.64, 0.36])
        elif kind == "image":
            block["path"] = local_file(root, block["path"])
        elif kind == "sources":
            if not block["items"]:
                raise ValueError("Source items require a label and URL")
            for label, url in block["items"]:
                target = urlsplit(url)
                if (
                    not label
                    or target.scheme not in ("https", "http")
                    or not target.netloc
                ):
                    raise ValueError(f"Invalid cited URL: {url}")
        elif kind not in role or not block.get("text", "").strip():
            raise ValueError(f"Unsupported or empty report block: {kind}")


def build_report(manifest, destination, overwrite=False):
    manifest = manifest.resolve()
    destination = destination.resolve()
    if destination.suffix.lower() != ".docx":
        raise ValueError("The output filename must end in .docx")
    if destination in (manifest, ASSET.resolve()) or (
        destination.exists() and not overwrite
    ):
        raise ValueError(f"Output file already exists: {destination}")
    data = json.loads(manifest.read_text(encoding="utf-8"))
    metadata = data["metadata"]
    if not metadata.get("company", "").strip():
        raise ValueError("metadata.company is required")
    pages = data["pages"]
    if not isinstance(pages, list) or not pages:
        raise ValueError("At least one page is required")
    for page in pages:
        prepare_blocks(page, manifest.parent)
    if metadata.get("logo"):
        metadata["logo"] = local_file(manifest.parent, metadata["logo"])

    document = Document(ASSET)
    document.element.body.clear_content()
    section = document.sections[0]
    width = (section.page_width - section.left_margin - section.right_margin) / 12700
    properties = document.core_properties
    properties.title = metadata["company"]
    properties.subject = metadata.get("report_kind", "公司深度研究")
    properties.author = metadata.get("author", "")
    properties.created = properties.modified = datetime.now(timezone.utc)
    properties.revision = 1
    document.settings.odd_and_even_pages_header_footer = False
    for existing in document.settings.element.xpath("./w:updateFields"):
        document.settings.element.remove(existing)
    document.settings.element.append(xml("updateFields", val="true"))

    set_headers(section, width, metadata)
    for index, page in enumerate(pages):
        if index == 1:
            next_section = document.add_section(WD_SECTION_START.NEW_PAGE)
            next_section.different_first_page_header_footer = False
            next_section.top_margin = Twips(1021)
            empty(document.paragraphs[-1])
        elif index > 1:
            break_paragraph = document.add_page_break()
            paragraph_format(break_paragraph, 1, 0, 0)
        run_blocks(document, page, width)
    destination.parent.mkdir(parents=True, exist_ok=True)
    document.save(destination)


if __name__ == "__main__":
    command = argparse.ArgumentParser(description=__doc__)
    command.add_argument("manifest", type=Path)
    command.add_argument("destination", type=Path)
    command.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace a report generated during this task.",
    )
    args = command.parse_args()
    try:
        build_report(args.manifest, args.destination, args.overwrite)
    except (KeyError, TypeError, ValueError, OSError) as exc:
        command.exit(2, f"{exc}\n")
