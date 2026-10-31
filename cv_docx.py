"""Word version of the CV: the Hero + tabs design from index.html, laid out as a one-page A4 .docx."""
from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Mm, Pt, RGBColor

# Same palette as the website. Calibri stands in for Nunito, which most computers don't have.
HERO = "2D5D8A"
ON_HERO = "FFFFFF"
HERO_SOFT = "D5E3F1"
ACCENT = "2F74B5"
ACCENT_SOFT = "E8F1FA"
SIDE_BG = "F3F7FB"
LINE = "C9D6E4"
INK = "26303D"
MUTED = "64707F"
FONT = "Calibri"

PAGE_W = Mm(210)
MAIN_W, SIDE_W = Mm(132), Mm(78)
PAD = Mm(12)
MAIN_TEXT_W = MAIN_W - PAD - Mm(5)


def _el(tag, **attrs):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn(f"w:{k}"), str(v))
    return el


def shade(cell, hex_):
    cell._tc.get_or_add_tcPr().append(_el("w:shd", val="clear", color="auto", fill=hex_))


def cell_margins(cell, top, left, bottom, right):
    mar = _el("w:tcMar")
    for edge, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
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
    table._tbl.tblPr.append(_el("w:tblW", w=sum(int(w.twips) for w in widths), type="dxa"))
    for col, w in zip(table.columns, widths):
        col.width = w
        for cell in col.cells:
            cell.width = w


def para(cell, first=False, after=0, before=0, line=1.12):
    p = cell.paragraphs[0] if first else cell.add_paragraph()
    pf = p.paragraph_format
    pf.space_after, pf.space_before, pf.line_spacing = Pt(after), Pt(before), line
    return p


def run(p, text, size=9.5, color=INK, bold=False, italic=False, tracking=None, fill=None, font=FONT):
    r = p.add_run(text)
    f = r.font
    f.name, f.size, f.bold, f.italic = font, Pt(size), bold, italic
    f.color.rgb = RGBColor.from_string(color)
    rpr = r._element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn("w:eastAsia"), font)
    if tracking:
        rpr.append(_el("w:spacing", val=int(tracking * 20)))
    if fill:
        rpr.append(_el("w:shd", val="clear", color="auto", fill=fill))
    return r


def link(p, url, text, size, color):
    r_id = p.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    h = _el("w:hyperlink")
    h.set(qn("r:id"), r_id)
    h.append(run(p, text, size=size, color=color)._element)
    p._p.append(h)


def bottom_rule(p, color=LINE):
    pbdr = _el("w:pBdr")
    pbdr.append(_el("w:bottom", val="single", sz=4, space=3, color=color))
    p._p.get_or_add_pPr().append(pbdr)


def timeline(p, indent=Mm(5)):
    """A left rule on consecutive paragraphs draws the timeline line, as on the website."""
    pbdr = _el("w:pBdr")
    pbdr.append(_el("w:left", val="single", sz=12, space=10, color=HERO_SOFT))
    p._p.get_or_add_pPr().append(pbdr)
    p.paragraph_format.left_indent = indent


def bullet(cell, text, size, color, indent, after=1):
    p = para(cell, after=after)
    p.paragraph_format.left_indent = indent
    p.paragraph_format.first_line_indent = -Mm(3.5)
    p.paragraph_format.tab_stops.add_tab_stop(indent)
    run(p, "•\t", size=size, color=ACCENT)
    run(p, text, size=size, color=color)
    return p


def heading(cell, text, before=10, first=False):
    p = para(cell, first=first, before=0 if first else before, after=4)
    run(p, text.upper(), size=9.5, color=INK, bold=True, tracking=1.6)


def side_heading(cell, text, first=False):
    p = para(cell, first=first, before=0 if first else 11, after=3)
    run(p, text.upper(), size=8.5, color=ACCENT, bold=True, tracking=1.6)


def write_docx(cv, path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = PAGE_W, Mm(297)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Mm(0)
    sec.header_distance = sec.footer_distance = Mm(0)
    doc.core_properties.title = f'{cv["name"]} — CV'
    doc.core_properties.author = cv["name"]
    c = cv["contact"]

    # Banner
    banner = doc.add_table(rows=1, cols=1)
    banner.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_borders(banner)
    fixed_layout(banner, (PAGE_W,))
    cell = banner.rows[0].cells[0]
    shade(cell, HERO)
    cell_margins(cell, Mm(9), PAD, Mm(7), PAD)
    p = para(cell, first=True, after=2); run(p, "CURRICULUM VITAE", 7.5, HERO_SOFT, bold=True, tracking=2.4)
    p = para(cell, after=3, line=1.0); run(p, cv["name"], 30, ON_HERO, bold=True, font="Georgia")
    p = para(cell, after=1); run(p, cv["current_role"], 13, ON_HERO, bold=True)
    p = para(cell); run(p, f'{cv["headline"]} · {cv["school"]}', 9.5, HERO_SOFT)

    # Body: main column + light side column, running to the bottom of the page
    body = doc.add_table(rows=1, cols=2)
    body.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_borders(body)
    fixed_layout(body, (MAIN_W, SIDE_W))
    row = body.rows[0]
    row.height, row.height_rule = Mm(236), WD_ROW_HEIGHT_RULE.AT_LEAST
    main, side = row.cells
    cell_margins(main, Mm(7), PAD, Mm(6), Mm(5))
    cell_margins(side, Mm(7), Mm(6), Mm(6), Mm(9))
    shade(side, SIDE_BG)

    heading(main, "About me", first=True)
    p = para(main, after=0, line=1.2); run(p, cv["profile"], 9.5, INK)

    heading(main, "Experience", before=12)
    for i, x in enumerate(cv["experience"]):
        p = para(main, before=0 if i == 0 else 7); timeline(p)
        p.paragraph_format.tab_stops.add_tab_stop(MAIN_TEXT_W, WD_TAB_ALIGNMENT.RIGHT)
        run(p, x["role"], 10.5, INK, bold=True)
        kind = (x.get("type") or "Full-time").upper()
        run(p, "\t")
        full = not x.get("type")
        run(p, f" {kind} ", 7, ON_HERO if full else ACCENT, bold=True, tracking=0.6, fill=ACCENT if full else ACCENT_SOFT)
        p = para(main); timeline(p); run(p, x["organisation"], 9, MUTED, italic=True)
        p = para(main, after=2); timeline(p); run(p, x["dates"], 8.5, ACCENT, bold=True)
        if x.get("bullets"):
            for b in x["bullets"]:
                timeline(bullet(main, b, 9, MUTED, indent=Mm(8.5)), indent=Mm(8.5))
        else:
            p = para(main); timeline(p); run(p, x.get("summary"), 9, MUTED)

    heading(main, "Leadership & Volunteering", before=12)
    n = len(cv["leadership"])
    vol = main.add_table(rows=1, cols=n)
    no_borders(vol)
    fixed_layout(vol, [Emu(int(MAIN_TEXT_W / n))] * n)
    for vc, x in zip(vol.rows[0].cells, cv["leadership"]):
        cell_margins(vc, 0, 0, 0, Mm(3))
        p = para(vc, first=True, after=1); run(p, x["organisation"], 9, INK, bold=True)
        p = para(vc); run(p, x["summary"], 8.5, MUTED)
    para(main)  # Word needs a paragraph after a nested table

    # Side column
    side_heading(side, "Contact", first=True)
    p = para(side, after=2); link(p, f'mailto:{c["email"]}', c["email"], 9, ACCENT)
    p = para(side, after=2); run(p, c["phone"], 9, ACCENT)
    p = para(side, after=0)
    if c.get("linkedin_url"):
        link(p, c["linkedin_url"], f'LinkedIn: {c["linkedin_label"]}', 9, ACCENT)
    else:
        run(p, f'LinkedIn: {c["linkedin_label"]}', 9, ACCENT)

    side_heading(side, "Skills")
    for i, s in enumerate(cv["skills"]):
        p = para(side, after=3, before=0 if i == 0 else 2); run(p, s, 8.5, INK)
        if i < len(cv["skills"]) - 1:
            bottom_rule(p)

    side_heading(side, "Languages")
    for i, l in enumerate(cv["languages"]):
        p = para(side, before=0 if i == 0 else 2); run(p, l["language"], 9, INK, bold=True)
        p = para(side, after=3); run(p, l["level"], 8.5, MUTED)
        if i < len(cv["languages"]) - 1:
            bottom_rule(p)

    side_heading(side, "Education")
    for i, x in enumerate(cv["education"]):
        p = para(side, before=0 if i == 0 else 3); run(p, x["degree"], 9, INK, bold=True)
        p = para(side); run(p, x["institution"], 8.5, MUTED)
        p = para(side, after=0 if x.get("note") else 3); run(p, x["dates"], 8.5, ACCENT, bold=True)
        if x.get("note"):
            p = para(side, after=3); run(p, x["note"], 8, MUTED)
        if i < len(cv["education"]) - 1:
            bottom_rule(p)

    # Word always keeps a paragraph after the last table; hide it so it can't start page 2.
    tail = doc.add_paragraph()
    tail.paragraph_format.space_after = tail.paragraph_format.space_before = Pt(0)
    tail.paragraph_format.line_spacing = Pt(1)
    run(tail, "", 1)
    mark = _el("w:rPr"); mark.append(_el("w:vanish")); mark.append(_el("w:sz", val=2))
    tail._p.get_or_add_pPr().append(mark)

    doc.save(path)
