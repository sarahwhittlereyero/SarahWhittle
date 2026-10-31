"""Word version of the CV: the same navy-sidebar layout as index.html, as a one-page A4 .docx."""
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

NAVY = "183A59"
SIDE_TEXT = "EEF3F8"
SIDE_MUTED = "B9C8D7"
INK = "2B2D30"
MUTED = "5A6068"
FONT = "Calibri"

SIDE_W, MAIN_W = Mm(72), Mm(138)
SIDE_PAD, MAIN_PAD = Mm(8), Mm(11)
MAIN_TEXT_W = MAIN_W - 2 * MAIN_PAD


def _el(tag, **attrs):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn(f"w:{k}"), str(v))
    return el


def shade(cell, hex_):
    cell._tc.get_or_add_tcPr().append(_el("w:shd", val="clear", color="auto", fill=hex_))


def cell_margins(cell, top, side, bottom):
    mar = _el("w:tcMar")
    for edge, v in (("top", top), ("left", side), ("bottom", bottom), ("right", side)):
        mar.append(_el(f"w:{edge}", w=int(v.twips if hasattr(v, "twips") else v), type="dxa"))
    cell._tc.get_or_add_tcPr().append(mar)


def no_borders(table):
    borders = _el("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_el(f"w:{edge}", val="nil"))
    table._tbl.tblPr.append(borders)


def fixed_layout(table, widths):
    table.autofit = False
    table._tbl.tblPr.append(_el("w:tblLayout", type="fixed"))
    for col, w in zip(table.columns, widths):
        col.width = w
        for cell in col.cells:
            cell.width = w


def para(cell, first=False, after=0, before=0, line=1.15):
    p = cell.paragraphs[0] if first else cell.add_paragraph()
    pf = p.paragraph_format
    pf.space_after, pf.space_before, pf.line_spacing = Pt(after), Pt(before), line
    return p


def run(p, text, size=9.5, color=INK, bold=False, italic=False, tracking=None):
    r = p.add_run(text)
    f = r.font
    f.name, f.size, f.bold, f.italic = FONT, Pt(size), bold, italic
    f.color.rgb = RGBColor.from_string(color)
    r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)
    if tracking:
        r._element.get_or_add_rPr().append(_el("w:spacing", val=int(tracking * 20)))
    return r


def link(p, url, text, size, color):
    r_id = p.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    h = _el("w:hyperlink")
    h.set(qn("r:id"), r_id)
    r = run(p, text, size=size, color=color)
    h.append(r._element)
    p._p.append(h)


def bullet(cell, text, size, color, indent=Mm(3.5), after=2):
    p = para(cell, after=after)
    p.paragraph_format.left_indent = indent
    p.paragraph_format.first_line_indent = -indent
    p.paragraph_format.tab_stops.add_tab_stop(indent)
    run(p, "•\t", size=size, color=color)
    run(p, text, size=size, color=color)
    return p


def timeline(p):
    """Left rule on consecutive paragraphs draws the experience timeline."""
    pbdr = _el("w:pBdr")
    pbdr.append(_el("w:left", val="single", sz=10, space=8, color=NAVY))
    p._p.get_or_add_pPr().append(pbdr)
    p.paragraph_format.left_indent = Mm(3)


def side_heading(cell, text, first=False):
    p = para(cell, first=first, before=0 if first else 14, after=6)
    run(p, text.upper(), size=11, color=SIDE_TEXT, bold=True, tracking=2.2)


def main_heading(cell, text, before=12):
    p = para(cell, before=before, after=5)
    run(p, text.upper(), size=11, color=NAVY, bold=True, tracking=2.2)


def write_docx(cv, path):
    doc = Document()
    sec = doc.sections[0]
    sec.start_type = WD_SECTION.NEW_PAGE
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Mm(0)
    sec.header_distance = sec.footer_distance = Mm(0)
    doc.core_properties.title = f'{cv["name"]} — CV'
    doc.core_properties.author = cv["name"]

    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_borders(table)
    fixed_layout(table, (SIDE_W, MAIN_W))
    row = table.rows[0]
    row.height, row.height_rule = Mm(296), WD_ROW_HEIGHT_RULE.AT_LEAST
    side, main = row.cells
    shade(side, NAVY)
    cell_margins(side, Mm(11), SIDE_PAD, Mm(8))
    cell_margins(main, Mm(11), MAIN_PAD, Mm(6))

    # Sidebar
    c = cv["contact"]
    side_heading(side, "Contact", first=True)
    p = para(side, after=3); run(p, c["phone"], 9, SIDE_TEXT)
    p = para(side, after=3); link(p, f'mailto:{c["email"]}', c["email"], 9, SIDE_TEXT)
    p = para(side, after=3)
    if c.get("linkedin_url"):
        run(p, "LinkedIn: ", 9, SIDE_MUTED); link(p, c["linkedin_url"], c["linkedin_label"], 9, SIDE_TEXT)
    else:
        run(p, f'LinkedIn: {c["linkedin_label"]}', 9, SIDE_TEXT)

    side_heading(side, "Education")
    for x in cv["education"]:
        p = para(side, after=1); run(p, x["institution"], 9, SIDE_TEXT, bold=True)
        p = para(side, after=0); run(p, x["degree"], 9, SIDE_TEXT)
        p = para(side, after=0 if x.get("note") else 7); run(p, x["dates"], 9, SIDE_MUTED)
        if x.get("note"):
            p = para(side, after=7); run(p, x["note"], 8, SIDE_MUTED)

    side_heading(side, "Skills")
    for s in cv["skills"]:
        bullet(side, s, 9, SIDE_TEXT)

    side_heading(side, "Languages")
    for l in cv["languages"]:
        p = bullet(side, l["language"], 9, SIDE_TEXT, after=0)
        p = para(side, after=4); p.paragraph_format.left_indent = Mm(3.5)
        run(p, l["level"], 8, SIDE_MUTED)

    # Main column
    p = para(main, first=True, line=1.0, after=6); run(p, cv["name"].upper(), 28, INK, bold=True)
    p = para(main, after=0); run(p, cv["headline"], 11.5, INK, tracking=0.6)
    p = para(main, after=4); run(p, cv["school"], 9.5, MUTED)
    p = para(main, after=0); run(p, cv["current_role"], 11.5, INK, tracking=0.6)

    main_heading(main, "About Me", before=16)
    p = para(main, after=0); p.alignment = 3  # justify
    run(p, cv["profile"], 9.5, INK)

    main_heading(main, "Experience")
    dates_tab = MAIN_TEXT_W
    for i, x in enumerate(cv["experience"]):
        p = para(main, before=0 if i == 0 else 6); timeline(p)
        run(p, x["role"], 10.5, INK, bold=True)
        p = para(main, after=2); timeline(p)
        p.paragraph_format.tab_stops.add_tab_stop(dates_tab, WD_TAB_ALIGNMENT.RIGHT)
        run(p, x["organisation"], 9, INK, italic=True)
        run(p, "\t" + " · ".join(filter(None, [x.get("type"), x.get("dates")])), 9, MUTED)
        if x.get("bullets"):
            for b in x["bullets"]:
                bp = bullet(main, b, 9, MUTED, after=1)
                timeline(bp)
                bp.paragraph_format.left_indent = Mm(6.5)
                bp.paragraph_format.first_line_indent = -Mm(3.5)
                bp.paragraph_format.tab_stops.add_tab_stop(Mm(6.5))
        else:
            p = para(main); timeline(p); run(p, x.get("summary"), 9, MUTED)

    main_heading(main, "Leadership & Volunteering")
    lead = main.add_table(rows=1, cols=len(cv["leadership"]))
    no_borders(lead)
    col_w = int(MAIN_TEXT_W / len(cv["leadership"]))
    fixed_layout(lead, [col_w] * len(cv["leadership"]))
    for cell, x in zip(lead.rows[0].cells, cv["leadership"]):
        cell_margins(cell, 0, Mm(1.5), 0)
        p = para(cell, first=True, after=2); run(p, x["organisation"], 9.5, INK, bold=True)
        p = para(cell); run(p, x["summary"], 8.5, MUTED)
    para(main, after=0)  # Word needs a paragraph after a nested table

    # Word always keeps a paragraph after the table; make it tiny so it doesn't start page 2.
    tail = doc.add_paragraph()
    tail.paragraph_format.space_after = tail.paragraph_format.space_before = Pt(0)
    tail.paragraph_format.line_spacing = Pt(1)
    run(tail, "", 1)
    mark = _el("w:rPr"); mark.append(_el("w:vanish")); mark.append(_el("w:sz", val=2))
    tail._p.get_or_add_pPr().append(mark)

    doc.save(path)
