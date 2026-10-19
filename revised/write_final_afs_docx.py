"""Editable Word copies of the signature financial statements.

The figures are the same as the PDF copies. Headings use Word styles, so the
navigation pane lists every section, and every amount sits in an ordinary table
cell that can be overwritten.
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Twips

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from revised.write_final_afs import (  # noqa: E402
    DIRECTORS,
    ISSUED,
    PAID_CAP,
    UNPAID,
    Y2022,
    YEARS,
    cash_lines,
    money,
    oi,
    ox,
)

NAVY = "1F4E79"
BLUE = "2F5597"
ACCENT = "D9E1F2"
ZEBRA = "F2F4F8"
PALE = "F7F9FC"
LINE = "D0D7E2"
WHITE = "FFFFFF"

# Usable width on A4 with 1.7 cm side margins.
PAGE_W = 17.6


def _rfonts(rPr, name="Calibri"):
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), name)


def set_run_font(run, name="Calibri", size=10.5, bold=False, color=None, italic=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._r.get_or_add_rPr()
    _rfonts(rPr, name)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def set_cell_margins(cell, top=40, bottom=40, left=60, right=60):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for edge, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = tcMar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            tcMar.append(node)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")


def set_valign(cell, val="center"):
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = tcPr.find(qn("w:vAlign"))
    if vAlign is None:
        vAlign = OxmlElement("w:vAlign")
        tcPr.append(vAlign)
    vAlign.set(qn("w:val"), val)


def prevent_row_split(row, header=False):
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))
    if header and trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))


def set_table_widths(table, widths):
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    total = sum(int(w.twips) for w in widths)
    tblW.set(qn("w:w"), str(total))
    tblW.set(qn("w:type"), "dxa")
    ind = tblPr.find(qn("w:tblInd"))
    if ind is None:
        ind = OxmlElement("w:tblInd")
        tblPr.append(ind)
    ind.set(qn("w:w"), "0")
    ind.set(qn("w:type"), "dxa")
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tblPr.addnext(grid)
    for child in list(grid):
        grid.remove(child)
    for w in widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(w.twips)))
        grid.append(gc)
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:w"), str(int(widths[i].twips)))
            tcW.set(qn("w:type"), "dxa")


def set_table_borders(table, color=LINE, sz="4"):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def no_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    tblPr.append(borders)


def para_border(paragraph, color=NAVY, sz="12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_field(paragraph, instruction, size=8, color="595959", bold=False):
    run = paragraph.add_run()
    set_run_font(run, size=size, color=color, bold=bold)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(separate)
    run._r.append(text)
    run._r.append(end)


def set_nowrap(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    if tcPr.find(qn("w:noWrap")) is None:
        tcPr.append(OxmlElement("w:noWrap"))


def cell_text(cell, text, *, size=9, bold=False, color="262626", align="left", fill=None, wrap=True):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }[align]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run("" if text is None else str(text))
    set_run_font(run, size=size, bold=bold, color=color)
    set_cell_margins(cell, top=36, bottom=36, left=70, right=70)
    set_valign(cell, "center")
    if fill:
        shade(cell, fill)
    if not wrap:
        set_nowrap(cell)
    return p


def add_table(doc, rows, widths, header=True, total_rows=None, section_rows=None, money_cols=None):
    total_rows = set(total_rows or [])
    section_rows = set(section_rows or [])
    money_cols = set(money_cols if money_cols is not None else range(1, len(rows[0])))
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.alignment = 1  # center
    set_table_widths(table, widths)
    set_table_borders(table)
    for r, row in enumerate(rows):
        prevent_row_split(table.rows[r], header=(r == 0 and header))
        for c, value in enumerate(row):
            if r == 0 and header:
                cell_text(table.cell(r, c), value, size=8, bold=True, color=WHITE, align="center", fill=NAVY)
            else:
                is_total = r in total_rows
                is_section = r in section_rows
                fill = ACCENT if is_total else (ZEBRA if is_section else None)
                align = "right" if c in money_cols else "left"
                color = NAVY if (is_total or is_section) else "262626"
                cell_text(
                    table.cell(r, c), value,
                    size=9, bold=is_total or is_section, color=color, align=align, fill=fill,
                    wrap=c not in money_cols,
                )
    if doc.paragraphs:
        doc.paragraphs[-1].paragraph_format.keep_with_next = True
    return table


def h1(doc, text):
    return doc.add_paragraph(text, style="Heading 1")


def h2(doc, text, *, tight=False, rule=False):
    p = doc.add_paragraph(text, style="Heading 2")
    if tight:
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
    if rule:
        para_border(p, BLUE, "12")
    return p


def body(doc, text, *, size=10.5, bold=False, space_after=6, center=False, color="262626"):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.08
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def body_lead(doc, lead, text, *, space_after=6):
    """A sentence whose opening label is bold, matching the signature notes."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.08
    r1 = p.add_run(lead)
    set_run_font(r1, size=10.5, bold=True, color=NAVY)
    r2 = p.add_run(text)
    set_run_font(r2, size=10.5, color="262626")
    return p


def banner(doc, text):
    table = doc.add_table(rows=1, cols=1)
    set_table_widths(table, [Cm(PAGE_W)])
    no_borders(table)
    cell = table.cell(0, 0)
    cell_text(cell, text, size=11, bold=True, color=WHITE, align="center", fill=NAVY)
    set_cell_margins(cell, top=90, bottom=90, left=80, right=80)
    return table


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.clear()
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    set_run_font(run, size=10.5, color="262626")
    return p


def signatures(doc, left, right):
    table = doc.add_table(rows=1, cols=2)
    set_table_widths(table, [Cm(PAGE_W / 2), Cm(PAGE_W / 2)])
    no_borders(table)
    for cell, lines in ((table.cell(0, 0), left), (table.cell(0, 1), right)):
        cell.text = ""
        set_cell_margins(cell, top=60, bottom=40, left=0, right=80)
        for i, (text, bold, size) in enumerate(lines):
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            run = p.add_run(text)
            set_run_font(run, size=size, bold=bold, color="262626")
    return table


def page_break(doc):
    doc.add_page_break()


def configure(doc, year):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.7)
    section.right_margin = Cm(1.7)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.6)
    section.header_distance = Cm(0.6)
    section.footer_distance = Cm(0.5)
    section.different_first_page_header_footer = True

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string("262626")
    _rfonts(normal.element.get_or_add_rPr())
    normal.paragraph_format.space_after = Pt(6)

    for style_name, size, color, before, after, bold in (
        ("Heading 1", 13, NAVY, 14, 2, True),
        ("Heading 2", 11, BLUE, 12, 2, True),
        ("Title", 20, NAVY, 0, 2, True),
    ):
        style = doc.styles[style_name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.color.rgb = RGBColor.from_string(color)
        style.font.italic = False
        _rfonts(style.element.get_or_add_rPr())
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.line_spacing = 1.05

    bullet_style = doc.styles["List Bullet"]
    bullet_style.font.name = "Calibri"
    bullet_style.font.size = Pt(10.5)
    _rfonts(bullet_style.element.get_or_add_rPr())

    def paint_header(header):
        p = header.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(PAGE_W), WD_TAB_ALIGNMENT.RIGHT)
        r1 = p.add_run("BAAY PROJECTS LIMITED (RC 1526224)")
        set_run_font(r1, size=9, bold=True, color=NAVY)
        r2 = p.add_run("\tAudited Financial Statements")
        set_run_font(r2, size=9, color="595959")
        para_border(p, NAVY, "10")

    def paint_footer(footer):
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(PAGE_W), WD_TAB_ALIGNMENT.RIGHT)
        r = p.add_run("BAAY PROJECTS LIMITED — ANNUAL REPORT AND FINANCIAL STATEMENTS")
        set_run_font(r, size=8, color="595959")
        p.add_run("\t")
        page_label = p.add_run("Page ")
        set_run_font(page_label, size=8, color="595959")
        add_field(p, " PAGE ", size=8)
        r3 = p.add_run(" of ")
        set_run_font(r3, size=8, color="595959")
        add_field(p, " NUMPAGES ", size=8)
        para_border(p, "D0D7E2", "6")
        # Move the border to the top of the footer paragraph.
        pPr = p._p.get_or_add_pPr()
        pBdr = pPr.find(qn("w:pBdr"))
        if pBdr is not None:
            bottom = pBdr.find(qn("w:bottom"))
            if bottom is not None:
                pBdr.remove(bottom)
            top = OxmlElement("w:top")
            top.set(qn("w:val"), "single")
            top.set(qn("w:sz"), "6")
            top.set(qn("w:space"), "1")
            top.set(qn("w:color"), "D0D7E2")
            pBdr.append(top)

    paint_header(section.header)
    paint_footer(section.footer)
    # Cover stays clear of the running header and page number, as in the signature PDF.
    section.first_page_header.paragraphs[0].text = ""
    section.first_page_footer.paragraphs[0].text = ""

    settings = doc.settings.element
    for tag, val in (("w:updateFields", "true"), ("w:autoHyphenation", "false")):
        node = settings.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            settings.append(node)
        node.set(qn("w:val"), val)

    core = doc.core_properties
    core.title = f"Baay Projects Limited — Audited financial statements {year}"
    core.subject = f"Annual report and audited financial statements for the year ended 31 December {year}"
    core.author = "Baay Projects Limited"
    core.category = "Financial statements"
    core.comments = (
        "Editable Word copy of the annual financial statements. "
        "Headings are Word styles. Amounts are ordinary table cells."
    )


def cv_map(year, c):
    if c:
        return lambda key: c.get(key, 0)
    opening = {
        "ppe": Y2022["ppe"], "inventory": 0, "rec_net": Y2022["receivables"],
        "other_rec": 0, "other_assets": 0, "related": 0, "cash": Y2022["cash"],
        "cl": 0, "co": 0, "other_liab": 0, "inter": 0, "tax_liab": 0,
        "re": Y2022["re"], "assets": Y2022["assets"], "liabilities": Y2022["liabilities"],
        "equity": Y2022["equity"], "payables": Y2022["payables"],
    }
    return lambda key: opening.get(key, 0)


def build(year, path):
    d = YEARS[year]
    comp = year - 1
    c = None if year == 2023 else YEARS[comp]
    cv = cv_map(year, c)
    doc = Document()
    configure(doc, year)

    # Cover
    body(doc, "BAAY PROJECTS LIMITED", size=22, bold=True, center=True, space_after=2, color=NAVY)
    body(doc, "(RC 1526224  •  TIN 21548976-0001)", size=12, bold=True, center=True, space_after=8, color=BLUE)
    rule = doc.add_paragraph()
    rule.paragraph_format.space_after = Pt(10)
    para_border(rule, NAVY, "18")
    body(doc, "ANNUAL REPORT AND AUDITED FINANCIAL STATEMENTS", size=16, bold=True, center=True, space_after=4, color=NAVY)
    body(doc, f"FOR THE YEAR ENDED 31 DECEMBER {year}", size=13, bold=True, center=True, space_after=2, color=BLUE)
    body(doc, f"(With Comparative Figures for the Year Ended 31 December {comp})", size=10, center=True, space_after=10, color="595959")
    banner(doc, "ANNUAL REPORT AND AUDITED FINANCIAL STATEMENTS")
    body(doc, "", space_after=8)

    info = [
        ["Registered Corporate Office:", "No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State"],
        ["Incorporation:", "18 September 2018  •  RC 1526224  •  Formerly Baay Degok Nig Ltd (name changed 28 June 2021)"],
        ["Company Secretary:", "Adegoke Mary Ayoboade"],
        ["Independent Auditors:", "Sanni Waheed & Co. (Chartered Accountants), 1st Floor, No. 29 Olonode Street, Alagomeji, Yaba, Lagos"],
        ["Principal Bankers:", "Providus Bank Plc  •  First Bank of Nigeria Limited  •  Sterling Bank Plc  •  GTBank"],
        ["Accounting Framework:", "International Financial Reporting Standards (IFRS) & CAMA 2020"],
    ]
    info_table = doc.add_table(rows=len(info), cols=2)
    set_table_widths(info_table, [Cm(5.6), Cm(PAGE_W - 5.6)])
    set_table_borders(info_table, BLUE, "6")
    for i, (label, value) in enumerate(info):
        prevent_row_split(info_table.rows[i])
        cell_text(info_table.cell(i, 0), label, size=9, bold=True, color=NAVY, fill=PALE)
        cell_text(info_table.cell(i, 1), value, size=9, fill=PALE)

    body(doc, "DIRECTORS", size=10, bold=True, space_after=3, color=NAVY)
    add_table(
        doc,
        [["No.", "Name of Director", "No.", "Name of Director"]]
        + [[str(a), an, str(b), bn] for a, an, b, bn in DIRECTORS],
        [Cm(1.3), Cm(7.5), Cm(1.3), Cm(7.5)],
        money_cols=set(),
    )
    body(doc, "SHAREHOLDERS", size=10, bold=True, space_after=3, color=NAVY)
    add_table(
        doc,
        [
            ["Name of shareholder", "Number of shares", "% of shareholding"],
            ["Adegoke Segun Babatunde", "80,000,000", "80%"],
            ["Adegoke Mary Ayoboade", "10,000,000", "10%"],
            ["Adegoke Denisa", "10,000,000", "10%"],
        ],
        [Cm(9.6), Cm(4.4), Cm(3.6)],
        money_cols={1, 2},
    )
    page_break(doc)

    # Directors' report
    h1(doc, "REPORT OF THE DIRECTORS")
    h2(doc, f"FOR THE YEAR ENDED 31 DECEMBER {year}", tight=True, rule=True)
    opening = doc.add_paragraph()
    opening.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    opening.paragraph_format.space_before = Pt(0)
    opening.paragraph_format.space_after = Pt(6)
    opening.paragraph_format.line_spacing = 1.08
    r = opening.add_run("The Directors submit their report together with the audited financial statements of ")
    set_run_font(r, size=10.5)
    r = opening.add_run("BAAY PROJECTS LIMITED")
    set_run_font(r, size=10.5, bold=True)
    r = opening.add_run(f" (\"the Company\") for the year ended 31 December {year}.")
    set_run_font(r, size=10.5)
    h2(doc, "1. Principal activities and corporate status")
    body(doc, (
        "The Company is a private company limited by shares, incorporated on 18 September 2018 as "
        "Baay Degok Nig Ltd and renamed Baay Projects Limited on 28 June 2021. Its registered office is "
        "No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State. The registered objects include integrated "
        "livestock farming, irrigation, cold-room operations, poultry, piggery and fishing. During the year "
        "the Company developed and sold land and housing, and carried on other trading."
    ))
    h2(doc, "2. Dividend")
    body(doc, (
        f"Dividends of ₦{d['dividends']:,.2f} were paid during the year. The Directors do not recommend a "
        "further dividend. The profit remaining has been carried forward."
    ))
    h2(doc, "3. Operating results")
    if c:
        comp_rev, comp_pbt, comp_pat, comp_div = c["revenue"], c["pbt"], c["pat"], c["dividends"]
    else:
        comp_rev, comp_pbt, comp_pat, comp_div = Y2022["revenue"], Y2022["pbt"], Y2022["pat"], 0
    add_table(
        doc,
        [
            ["", f"{year}  ₦", f"{comp}  ₦"],
            ["Revenue", money(d["revenue"]), money(comp_rev)],
            ["Profit before taxation", money(d["pbt"]), money(comp_pbt)],
            ["Profit for the year", money(d["pat"]), money(comp_pat)],
            ["Dividends paid", money(d["dividends"]), money(comp_div)],
        ],
        [Cm(9.6), Cm(4.0), Cm(4.0)],
        total_rows={3},
    )
    h2(doc, "4. Board of Directors")
    body(doc, "The directors who served are those named in the Corporate Affairs Commission status report of 26 August 2026.")
    add_table(
        doc,
        [["No.", "Name of director", "No.", "Name of director"]]
        + [[str(a), an, str(b), bn] for a, an, b, bn in DIRECTORS],
        [Cm(1.3), Cm(7.5), Cm(1.3), Cm(7.5)],
        money_cols=set(),
    )
    h2(doc, "5. Shareholders")
    add_table(
        doc,
        [
            ["Name of shareholder", "Number of shares", "% of shareholding"],
            ["Adegoke Segun Babatunde", "80,000,000", "80%"],
            ["Adegoke Mary Ayoboade", "10,000,000", "10%"],
            ["Adegoke Denisa", "10,000,000", "10%"],
        ],
        [Cm(9.6), Cm(4.4), Cm(3.6)],
        money_cols={1, 2},
    )
    body(doc, (
        "The interests above are the holdings recorded at the Corporate Affairs Commission in the issued "
        "capital of ₦100,000,000. Note 18 sets out the amount recognised in these financial statements."
    ))
    h2(doc, "6. Independent auditors")
    body(doc, (
        "Sanni Waheed & Co. (Chartered Accountants), of 1st Floor, No. 29 Olonode Street, Alagomeji, Yaba, Lagos, "
        "are the auditors. They have indicated their willingness to continue in office in accordance with "
        "section 401 of the Companies and Allied Matters Act 2020."
    ))
    body(doc, "", space_after=8)
    signatures(
        doc,
        [("BY ORDER OF THE BOARD", True, 10), ("", False, 10), ("______________________________", False, 11),
         ("Adegoke Mary Ayoboade", True, 10.5), ("Company Secretary", False, 10)],
        [("FOR AND ON BEHALF OF THE BOARD", True, 10), ("", False, 10), ("______________________________", False, 11),
         ("Adegoke Segun Babatunde", True, 10.5), ("Managing Director", False, 10)],
    )
    page_break(doc)

    # Responsibilities
    h1(doc, "STATEMENT OF DIRECTORS' RESPONSIBILITIES")
    h2(doc, f"IN RELATION TO THE FINANCIAL STATEMENTS FOR THE YEAR ENDED 31 DECEMBER {year}", tight=True, rule=True)
    body(doc, (
        "The Companies and Allied Matters Act 2020 and the Financial Reporting Council of Nigeria Act require "
        "the Directors to prepare financial statements for each financial year that give a true and fair view of "
        "the state of affairs of BAAY PROJECTS LIMITED at the end of the year and of its profit or loss and "
        "cash flows for the year then ended."
    ))
    body(doc, "In preparing these financial statements, the Directors are required to:")
    for item in [
        "select suitable accounting policies in compliance with IFRS Accounting Standards and apply them consistently;",
        "make judgements and accounting estimates that are reasonable and prudent;",
        "state whether applicable IFRS Accounting Standards have been followed, subject to any material departures disclosed and explained in the financial statements; and",
        "prepare the financial statements on the going concern basis unless it is inappropriate to presume that the Company will continue in business.",
    ]:
        bullet(doc, item)
    body(doc, (
        "The Directors are responsible for keeping proper accounting records that disclose with reasonable "
        "accuracy at any time the financial position of the Company and enable them to ensure that the "
        "financial statements comply with the Companies and Allied Matters Act 2020 and IFRS Accounting "
        "Standards. They are also responsible for safeguarding the assets of the Company and for taking "
        "reasonable steps for the prevention and detection of fraud and other irregularities."
    ))
    body(doc, (
        f"The financial statements for the year ended 31 December {year} were approved by the Board of "
        "Directors and signed on its behalf by:"
    ))
    body(doc, "", space_after=16)
    signatures(
        doc,
        [("______________________________", False, 11), ("Adegoke Segun Babatunde", True, 10.5), ("Managing Director", False, 10)],
        [("______________________________", False, 11), ("Owolabi Charles Oluwatobi", True, 10.5), ("Director", False, 10)],
    )
    page_break(doc)

    # Auditor
    h1(doc, "INDEPENDENT AUDITOR'S REPORT")
    h2(doc, "TO THE MEMBERS OF BAAY PROJECTS LIMITED (RC 1526224)", tight=True, rule=True)
    h2(doc, "Qualified opinion")
    body(doc, (
        f"We have audited the financial statements of Baay Projects Limited (the Company), which comprise the "
        f"statement of financial position as at 31 December {year}, and the statement of profit or loss and "
        f"other comprehensive income, statement of changes in equity and statement of cash flows for the year "
        f"then ended, and notes to the financial statements, including a summary of significant accounting policies."
    ))
    body(doc, (
        f"In our opinion, except for the possible effects of the matter described in the Basis for Qualified "
        f"Opinion section of our report, the accompanying financial statements give a true and fair view of the "
        f"financial position of the Company as at 31 December {year}, and of its financial performance and its "
        f"cash flows for the year then ended, in accordance with IFRS Accounting Standards and the requirements "
        f"of the Companies and Allied Matters Act 2020 and the Financial Reporting Council of Nigeria Act."
    ))
    h2(doc, "Basis for qualified opinion")
    body(doc, (
        f"Other assets of ₦{d['payments_unid']:,.2f} and other liabilities of ₦{d['other_liab']:,.2f} are "
        f"amounts whose nature has not been identified. We were unable to determine whether any adjustment to "
        f"profit, inventories, receivables or liabilities is required."
    ))
    body(doc, (
        "We conducted our audit in accordance with International Standards on Auditing. Our responsibilities "
        "under those standards are described in the Auditor's Responsibilities for the Audit of the Financial "
        "Statements section of our report. We are independent of the Company in accordance with the "
        "International Ethics Standards Board for Accountants' International Code of Ethics for Professional "
        "Accountants (including International Independence Standards), and we have fulfilled our other ethical "
        "responsibilities in accordance with these requirements. We believe that the audit evidence we have "
        "obtained is sufficient and appropriate to provide a basis for our qualified opinion."
    ))
    h2(doc, "Key audit matters")
    body(doc, (
        "Revenue and contract liabilities. Revenue from land is recognised when control of an identified plot "
        "has transferred. Revenue from housing is recognised when handover is evidenced. Receipts before that "
        "point are contract liabilities. We examined the status of the contracts and the timing of recognition."
    ))
    body(doc, (
        "Inventories. Land, housing stock and development expenditure are carried at cost and released to "
        "cost of sales as revenue is recognised. We tested the identification of those costs and the release "
        "to profit or loss."
    ))
    h2(doc, "Other information")
    body(doc, (
        "The Directors are responsible for the other information, which comprises the report of the directors. "
        "Our opinion on the financial statements does not cover the other information and we do not express "
        "any form of assurance conclusion thereon."
    ))
    h2(doc, "Report on other legal and regulatory requirements")
    body(doc, (
        "As required by the Companies and Allied Matters Act 2020, except for the possible effects of the "
        "matter described in the Basis for Qualified Opinion section, we confirm that:"
    ))
    body(doc, "i. we have obtained all the information and explanations which, to the best of our knowledge and belief, were necessary for the purpose of our audit;", space_after=2)
    body(doc, "ii. proper books of account have been kept by the Company, so far as appears from our examination of those books; and", space_after=2)
    body(doc, "iii. the Company's statement of financial position and statement of profit or loss and other comprehensive income are in agreement with the books of account.", space_after=6)
    body(doc, "", space_after=8)
    body(doc, "Sanni Waheed & Co.", size=11, bold=True, space_after=0)
    body(doc, "Chartered Accountants", size=10.5, space_after=0)
    body(doc, "1st Floor, No. 29 Olonode Street, Alagomeji, Yaba, Lagos", size=10.5, space_after=10)
    body(doc, "________________________________", size=11, space_after=0)
    body(doc, "Sanni Waheed, FCA", size=11, bold=True, space_after=0)
    body(doc, "Engagement Partner", size=10.5, space_after=0)
    body(doc, "FRC/2016/ICAN/2016/00000013886", size=10.5, space_after=0)
    body(doc, f"Date: 28 April {year + 1}", size=10.5, space_after=0)
    page_break(doc)

    # Statement of financial position
    h1(doc, "STATEMENT OF FINANCIAL POSITION")
    h2(doc, f"AS AT 31 DECEMBER {year}", tight=True, rule=True)
    issued_c = ISSUED if c else Y2022["share"]
    unpaid_c = UNPAID if c else 0
    sfp = [
        ["", "Notes", f"31 Dec {year}  ₦", f"31 Dec {comp}  ₦"],
        ["NON-CURRENT ASSETS", "", "", ""],
        ["Property, plant and equipment", "7", money(d["ppe"]), money(cv("ppe"))],
        ["CURRENT ASSETS", "", "", ""],
        ["Inventories", "8", money(d["inventory"]), money(cv("inventory"))],
        ["Trade receivables", "10", money(d["rec_net"]), money(cv("rec_net"))],
        ["Other receivables", "11", money(d["other_rec"]), money(cv("other_rec"))],
        ["Other assets", "12", money(d["other_assets"]), money(cv("other_assets"))],
        ["Amounts due from related parties", "17", money(d["related"]), money(cv("related"))],
        ["Cash and cash equivalents", "13", money(d["cash"]), money(cv("cash"))],
        ["Total assets", "", money(d["assets"]), money(cv("assets"))],
        ["EQUITY", "", "", ""],
        ["Issued share capital", "18", money(ISSUED), money(issued_c)],
        ["Amount due on issued shares", "18", money(-UNPAID), money(-unpaid_c)],
        ["Retained earnings", "", money(d["re"]), money(cv("re"))],
        ["Total equity", "", money(d["equity"]), money(cv("equity"))],
        ["LIABILITIES", "", "", ""],
        ["Contract liabilities", "9", money(d["cl"]), money(cv("cl"))],
        ["Trade and other payables", "14", "—", money(cv("payables"))],
        ["Co-ownership liabilities", "14", money(d["co"]), money(cv("co"))],
        ["Other liabilities", "12", money(d["other_liab"]), money(cv("other_liab"))],
        ["Amounts due on the Company's accounts", "15", money(d["inter"]), money(cv("inter"))],
        ["Current tax liabilities", "16", money(d["tax_liab"]), money(cv("tax_liab"))],
        ["Total liabilities", "", money(d["liabilities"]), money(cv("liabilities"))],
        ["Total equity and liabilities", "", money(d["assets"]), money(cv("assets"))],
    ]
    add_table(
        doc, sfp, [Cm(9.2), Cm(1.6), Cm(3.4), Cm(3.4)],
        section_rows={1, 3, 11, 16}, total_rows={10, 15, 23, 24}, money_cols={2, 3},
    )
    body(doc, "The financial statements were approved by the Board of Directors and signed on its behalf by:", space_after=12)
    signatures(
        doc,
        [("______________________________", False, 11), ("Adegoke Segun Babatunde", True, 10.5), ("Managing Director", False, 10)],
        [("______________________________", False, 11), ("Owolabi Charles Oluwatobi", True, 10.5), ("Director", False, 10)],
    )
    page_break(doc)

    # Profit or loss and equity
    h1(doc, "STATEMENT OF PROFIT OR LOSS AND OTHER COMPREHENSIVE INCOME")
    h2(doc, f"FOR THE YEAR ENDED 31 DECEMBER {year}", tight=True, rule=True)
    gp = d["revenue"] - d["cos"]
    op = gp + oi(d) - ox(d)
    if c:
        gp_c = c["revenue"] - c["cos"]
        oi_c, ox_c = oi(c), ox(c)
        op_c = gp_c + oi_c - ox_c
        rev_c, cos_c, fin_c = c["revenue"], c["cos"], c["finance"]
        pbt_c, tax_c, pat_c = c["pbt"], c["tax"], c["pat"]
    else:
        gp_c, oi_c, ox_c, op_c = Y2022["revenue"], 0, Y2022["opex"], Y2022["pat"]
        rev_c, cos_c, fin_c = Y2022["revenue"], 0, 0
        pbt_c, tax_c, pat_c = Y2022["pbt"], 0, Y2022["pat"]
    add_table(
        doc,
        [
            ["", "Notes", f"{year}  ₦", f"{comp}  ₦"],
            ["Revenue", "3", money(d["revenue"]), money(rev_c)],
            ["Cost of sales", "4", money(-d["cos"]), money(-cos_c)],
            ["Gross profit", "", money(gp), money(gp_c)],
            ["Other income", "3", money(oi(d)), money(oi_c)],
            ["Administrative and operating expenses", "5", money(-ox(d)), money(-ox_c)],
            ["Operating profit", "", money(op), money(op_c)],
            ["Finance costs", "6", money(-d["finance"]), money(-fin_c)],
            ["Profit before taxation", "", money(d["pbt"]), money(pbt_c)],
            ["Income tax expense", "16", money(-d["tax"]), money(-tax_c)],
            ["Profit for the year", "", money(d["pat"]), money(pat_c)],
            ["Other comprehensive income", "", "—", "—"],
            ["Total comprehensive income", "", money(d["pat"]), money(pat_c)],
        ],
        [Cm(9.2), Cm(1.6), Cm(3.4), Cm(3.4)],
        section_rows={3, 6}, total_rows={8, 10, 12}, money_cols={2, 3},
    )
    body(doc, "There was no other comprehensive income.", space_after=10)
    h1(doc, "STATEMENT OF CHANGES IN EQUITY")
    h2(doc, f"FOR THE YEAR ENDED 31 DECEMBER {year}", tight=True, rule=True)
    if c:
        open_issued, open_unpaid, open_re, open_eq = ISSUED, UNPAID, c["re"], c["equity"]
    else:
        open_issued, open_unpaid, open_re, open_eq = Y2022["share"], 0, Y2022["re"], Y2022["equity"]
    soce = [
        ["", "Issued capital  ₦", "Unpaid calls  ₦", "Retained earnings  ₦", "Total  ₦"],
        [f"At 1 January {year}", money(open_issued), money(-open_unpaid), money(open_re), money(open_eq)],
    ]
    if not c:
        soce.append(["Share capital issued, not paid", money(UNPAID), money(-UNPAID), "—", "—"])
        other_eq = d["re"] - (open_re + d["pat"] - d["dividends"])
        soce.append(["Profit for the year", "—", "—", money(d["pat"]), money(d["pat"])])
        soce.append(["Dividends paid", "—", "—", money(-d["dividends"]), money(-d["dividends"])])
        soce.append(["Other movements", "—", "—", money(other_eq), money(other_eq)])
    else:
        soce.append(["Profit for the year", "—", "—", money(d["pat"]), money(d["pat"])])
        soce.append(["Dividends paid", "—", "—", money(-d["dividends"]), money(-d["dividends"])])
    soce.append([f"At 31 December {year}", money(ISSUED), money(-UNPAID), money(d["re"]), money(d["equity"])])
    add_table(
        doc, soce, [Cm(4.6), Cm(3.3), Cm(3.1), Cm(3.5), Cm(3.1)],
        total_rows={len(soce) - 1}, money_cols={1, 2, 3, 4},
    )
    if not c:
        body(doc, (
            "Other movements are recognised in equity so that the closing balance equals assets less liabilities. "
            "They are not included in profit for the year."
        ))
    page_break(doc)

    # Cash flow
    h1(doc, "STATEMENT OF CASH FLOWS")
    h2(doc, f"FOR THE YEAR ENDED 31 DECEMBER {year}", tight=True, rule=True)
    cur = cash_lines(year)
    comp_map = {}
    if year > 2023:
        comp_map = {name: val for name, val, _ in cash_lines(comp)}
    else:
        comp_map = {
            "Net cash from operating activities": Y2022["cf_op"],
            "Net cash from investing activities": 0,
            "Net cash from financing activities": 0,
            "Net increase / (decrease) in cash": Y2022["cf_net"],
            "Cash and cash equivalents at 1 January": Y2022["cash_open"],
            "Cash and cash equivalents at 31 December": Y2022["cash_close"],
        }
    cf_rows = [["", f"{year}  ₦", f"{comp}  ₦"]]
    section_rows, total_rows = set(), set()
    for name, val, is_total in cur:
        if val is None:
            cf_rows.append([name, "", ""])
            section_rows.add(len(cf_rows) - 1)
        else:
            shown = comp_map.get(name, None)
            cf_rows.append([name, money(val), money(shown) if shown is not None else "—"])
            if is_total:
                total_rows.add(len(cf_rows) - 1)
    add_table(
        doc, cf_rows, [Cm(10.2), Cm(3.7), Cm(3.7)],
        section_rows=section_rows, total_rows=total_rows, money_cols={1, 2},
    )
    note = f"Cash at 31 December {year} is the balance held at the bank."
    if year == 2023:
        note += (
            " Other cash movements are included so that the closing balance equals that bank balance. "
            "Comparative totals are the amounts reported for the year ended 31 December 2022."
        )
    body(doc, note)
    page_break(doc)

    # Notes
    h1(doc, "NOTES TO THE FINANCIAL STATEMENTS")
    h2(doc, f"FOR THE YEAR ENDED 31 DECEMBER {year}", tight=True, rule=True)
    write_notes(doc, year, d, c)
    body(doc, "", space_after=8)
    h1(doc, "STATEMENT OF VALUE ADDED")
    h2(doc, f"FOR THE YEAR ENDED 31 DECEMBER {year}", tight=True, rule=True)
    employees = d["staff"] + d["director_rem"]
    bought = d["cos"] + d["refunds"] + d["commission"] + d["rent"] + d["professional"] + d["admin"] + d["bank_charges"]
    created = d["revenue"] + oi(d) - bought
    retained = d["depreciation"] + d["ecl"] + d["pat"]
    distributed = employees + d["tax"] + d["finance"] + retained
    if abs(created - distributed) > 1:
        raise SystemExit(f"Value added does not tie for {year}")
    def pct(x):
        return f"{x / created * 100:.1f}%"
    add_table(
        doc,
        [
            ["", f"{year}  ₦", "%"],
            ["Revenue and other income", money(d["revenue"] + oi(d)), ""],
            ["Bought-in materials and services", money(-bought), ""],
            ["Value added", money(created), "100.0%"],
            ["Applied as follows", "", ""],
            ["To employees", money(employees), pct(employees)],
            ["To government — taxation", money(d["tax"]), pct(d["tax"])],
            ["To providers of finance", money(d["finance"]), pct(d["finance"])],
            ["Retained in the business", money(retained), pct(retained)],
            ["Value applied", money(distributed), "100.0%"],
        ],
        [Cm(11.0), Cm(4.0), Cm(2.6)],
        section_rows={4}, total_rows={3, 9}, money_cols={1, 2},
    )
    body(doc, (
        "Value added is revenue and other income less bought-in materials and services. It is applied to "
        "employees, government, providers of finance, and the amount retained in the business."
    ))

    doc.save(path)
    print("wrote", path)


def note_table(doc, rows, total_rows=None, section_rows=None):
    cols = len(rows[0])
    if cols == 3:
        widths = [Cm(9.6), Cm(4.0), Cm(4.0)]
        money_cols = {1, 2}
    else:
        widths = [Cm(PAGE_W)]
        money_cols = set()
    return add_table(doc, rows, widths, total_rows=total_rows, section_rows=section_rows, money_cols=money_cols)


def write_notes(doc, year, d, c):
    comp = year - 1
    h2(doc, "1. General information")
    body(doc, (
        "Baay Projects Limited was incorporated in Nigeria on 18 September 2018 as a private company limited "
        "by shares (RC 1526224). It was formerly named Baay Degok Nig Ltd. The name was changed on 28 June 2021. "
        "The registered office is No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State. The tax identification "
        "number is 21548976-0001."
    ))
    body(doc, (
        "The financial statements are for the Company alone. The directors, the company secretary and the "
        "shareholders are those set out in the report of the directors."
    ))
    h2(doc, "2. Basis of preparation and accounting policies")
    body(doc, (
        f"The financial statements have been prepared in accordance with IFRS Accounting Standards and the "
        f"Companies and Allied Matters Act 2020. They are presented in Nigerian Naira under the historical cost "
        f"convention. The Company held cash of ₦{d['cash']:,.2f} at 31 December {year} and continued to receive "
        f"amounts from customers. The Directors have prepared the financial statements on the going concern basis."
    ))
    if not c:
        body(doc, "Comparative figures for the year ended 31 December 2022 are the amounts reported for that year. Where a line was not reported, no comparative is shown.")
    body_lead(doc, "Revenue. ", "Revenue from the sale of land is recognised when control of an identified plot has transferred. Housing revenue is recognised when handover is evidenced and the contract is substantially paid. Receipts before that point are contract liabilities. Where the amount still unpaid exceeds the balance on the customer register, revenue and the receivable are limited to that balance.")
    body_lead(doc, "Inventories. ", "Land, housing stock and development expenditure are measured at cost. Cost is released to cost of sales as the related revenue is recognised.")
    body_lead(doc, "Financial instruments. ", "Receivables are carried at cost less a loss allowance of 5 per cent. Cash is the balance held at the bank.")
    body_lead(doc, "Property, plant and equipment. ", "Items are carried at cost less accumulated depreciation. Depreciation is 20 per cent a year on a straight-line basis, with a full year in the year of purchase. The same rate is applied to plant and machinery, office equipment and computer equipment.")
    body_lead(
        doc,
        "Income tax. ",
        f"Current tax is provided under the law applicable to the year. The Company is a {d['size']} "
        f"company. Companies income tax is provided at {d['rate']:.0%} of assessable profit. Tertiary education tax "
        f"is 3 per cent. The police trust fund levy is 0.005 per cent of profit before tax. A National Agency for "
        f"Science and Engineering Infrastructure levy is accrued for a large company. The Nigeria Tax Act 2025 "
        f"applies to periods beginning on or after 1 January 2026 and has not been applied. Deferred tax has not been recognised.",
    )
    body_lead(doc, "Amounts not yet identified. ", "A receipt or payment whose nature is not identified is not included in profit. Receipts are carried as other liabilities. Payments are carried as other assets.")

    h2(doc, "3. Revenue and other income")
    note_table(doc, [
        ["Revenue", f"{year}  ₦", f"{comp}  ₦"],
        ["Land", money(d["rev_land"]), money(c["rev_land"]) if c else ""],
        ["Housing", money(d["rev_hous"]), money(c["rev_hous"]) if c else ""],
        ["Revenue", money(d["revenue"]), money(c["revenue"] if c else Y2022["revenue"])],
    ], total_rows={3})
    note_table(doc, [
        ["Other income", f"{year}  ₦", f"{comp}  ₦"],
        ["Trading", money(d["other_trading"]), money(c["other_trading"]) if c else ""],
        ["Service income", money(d["other_service"]), money(c["other_service"]) if c else ""],
        ["Sundry receipts", money(d["other_refunds"] + d["bank_credit"]), money(c["other_refunds"] + c["bank_credit"]) if c else ""],
        ["Other income", money(oi(d)), money(oi(c) if c else 0)],
    ], total_rows={4})
    if not c:
        body(doc, "Comparative revenue is the turnover reported for the year ended 31 December 2022.")

    h2(doc, "4. Cost of sales")
    note_table(doc, [
        ["", f"{year}  ₦", f"{comp}  ₦"],
        ["Land, housing and development", money(d["cos"] - d["cos_trading"]), money(c["cos"] - c["cos_trading"]) if c else ""],
        ["Trading purchases", money(d["cos_trading"]), money(c["cos_trading"]) if c else ""],
        ["Cost of sales", money(d["cos"]), money(c["cos"]) if c else money(0)],
    ], total_rows={3})

    h2(doc, "5. Administrative and operating expenses")
    def ox_line(key):
        return money(c[key]) if c else ""
    note_table(doc, [
        ["", f"{year}  ₦", f"{comp}  ₦"],
        ["Customer refunds", money(d["refunds"]), ox_line("refunds")],
        ["Staff costs", money(d["staff"]), ox_line("staff")],
        ["Director remuneration", money(d["director_rem"]), ox_line("director_rem")],
        ["Selling commissions", money(d["commission"]), ox_line("commission")],
        ["Rent", money(d["rent"]), ox_line("rent")],
        ["Professional fees", money(d["professional"]), ox_line("professional")],
        ["Administrative expenses", money(d["admin"]), ox_line("admin")],
        ["Bank charges", money(d["bank_charges"]), ox_line("bank_charges")],
        ["Depreciation", money(d["depreciation"]), money(c["depreciation"]) if c else ""],
        ["Impairment of receivables", money(d["ecl"]), ox_line("ecl")],
        ["Total", money(ox(d)), money(ox(c) if c else Y2022["opex"])],
    ], total_rows={11})
    if not c:
        body(doc, "The comparative total is the amount reported for the year ended 31 December 2022. It was not analysed on the same lines.")

    h2(doc, "6. Finance costs")
    extra = f" (₦{c['finance']:,.2f} in {comp})" if c else ""
    body(doc, f"Finance costs of ₦{d['finance']:,.2f}{extra} are interest and similar charges.")

    h2(doc, "7. Property, plant and equipment")
    note_table(doc, [
        ["", f"{year}  ₦", f"{comp}  ₦"],
        ["Cost", money(d["ppe_cost"]), money(c["ppe_cost"]) if c else ""],
        ["Accumulated depreciation", money(-d["accum_dep"]), money(-c["accum_dep"]) if c else ""],
        ["Carrying amount", money(d["ppe"]), money(c["ppe"] if c else Y2022["ppe"])],
    ], total_rows={3})
    body(doc, f"Depreciation for the year is ₦{d['depreciation']:,.2f}.")

    h2(doc, "8. Inventories")
    body(doc, f"Inventories of ₦{d['inventory']:,.2f} are land, housing units and development expenditure, less the cost released to cost of sales.")
    h2(doc, "9. Contract liabilities")
    body(doc, f"Contract liabilities of ₦{d['cl']:,.2f} are amounts received from customers for performance that has not yet been completed. The liability becomes revenue when control transfers.")

    h2(doc, "10. Trade receivables")
    note_table(doc, [
        ["", f"{year}  ₦", f"{comp}  ₦"],
        ["Gross receivables", money(d["rec_gross"]), money(c["rec_gross"]) if c else money(Y2022["receivables"])],
        ["Loss allowance", money(-d["rec_ecl"]), money(-c["rec_ecl"]) if c else ""],
        ["Net receivables", money(d["rec_net"]), money(c["rec_net"]) if c else money(Y2022["receivables"])],
    ], total_rows={3})
    body(doc, "A receivable is recognised for the unpaid balance of a contract whose revenue has been recognised, and not above the balance on the customer register. The loss allowance is 5 per cent.")

    h2(doc, "11. Other receivables")
    body(doc, f"Other receivables of ₦{d['other_rec']:,.2f} are amounts recorded as due, other than trade receivables on contracts whose revenue has been recognised.")

    h2(doc, "12. Other assets and other liabilities")
    note_table(doc, [
        ["", f"{year}  ₦", f"{comp}  ₦"],
        ["Payments not yet identified", money(d["payments_unid"]), money(c["payments_unid"]) if c else ""],
        ["Foreign-currency and other advances", money(d["fx"]), money(c["fx"]) if c else ""],
        ["Loan repayments in excess of proceeds identified", money(d["loan_excess"]), money(c["loan_excess"]) if c else ""],
        ["Other assets", money(d["other_assets"]), money(c["other_assets"]) if c else ""],
        ["Receipts not yet identified", money(d["other_liab"]), money(c["other_liab"]) if c else ""],
    ], total_rows={4})
    body(doc, "These amounts have not been included in profit. Profit for the year is stated before any later identification of them.")

    h2(doc, "13. Cash and cash equivalents")
    note_table(doc, [
        ["", f"{year}  ₦", f"{comp}  ₦"],
        ["Balance held with banks", money(d["cash"]), money(c["cash"] if c else Y2022["cash"])],
        ["Cash", "—", "—"],
        ["Total", money(d["cash"]), money(c["cash"] if c else Y2022["cash"])],
    ], total_rows={3})

    h2(doc, "14. Trade and other payables, and co-ownership")
    text = (
        f"Co-ownership liabilities of ₦{d['co']:,.2f} are investment receipts that are not sales of land or housing, "
        "less returns paid. They are not revenue."
    )
    if not c:
        text += f" Trade and other payables reported at 31 December {comp} were ₦{Y2022['payables']:,.2f}."
    body(doc, text)

    h2(doc, "15. Amounts due on the Company's accounts")
    body(doc, f"₦{d['inter']:,.2f} is the net amount due on accounts of the Company other than the balances included in cash. It is carried as a liability.")

    h2(doc, "16. Taxation")
    note_table(doc, [
        ["", f"{year}  ₦", f"{comp}  ₦"],
        ["Companies income tax", money(d["cit"]), money(c["cit"]) if c else "—"],
        ["Tertiary education tax", money(d["tet"]), money(c["tet"]) if c else "—"],
        ["Police trust fund levy", money(d["ptf"]), money(c["ptf"]) if c else "—"],
        ["NASENI levy", money(d["naseni"]), money(c["naseni"]) if c else "—"],
        ["Income tax expense", money(d["tax"]), money(c["tax"]) if c else "—"],
        ["Current tax liability", money(d["tax_liab"]), money(c["tax_liab"]) if c else "—"],
    ], total_rows={5})
    body(doc, f"The Company is a {d['size']} company. The liability is the charge for each year less tax payments identified. It is not an agreed assessment.")

    h2(doc, "17. Related parties")
    body(doc, (
        f"Amounts due from related parties at 31 December {year} are ₦{d['related']:,.2f}. The balance is the net "
        f"of amounts received from and paid to the directors, principally Adegoke Segun Babatunde, after director "
        f"remuneration of ₦{d['director_rem']:,.2f} has been charged to profit and after dividends of "
        f"₦{d['dividends']:,.2f} have been treated as distributions. A debit balance is an amount due to the Company."
    ))
    body(doc, "Key management personnel are the directors. Remuneration charged in profit is the director remuneration in Note 5.")

    h2(doc, "18. Share capital")
    note_table(doc, [
        ["", f"{year}  ₦", f"{comp}  ₦"],
        ["Issued ordinary shares of ₦1 each", money(ISSUED), money(ISSUED if c else Y2022["share"])],
        ["Amount due on issued shares", money(-UNPAID), money(-UNPAID if c else 0)],
        ["Share capital recognised", money(PAID_CAP), money(PAID_CAP)],
    ], total_rows={3})
    body(doc, (
        "The Corporate Affairs Commission status report of 26 August 2026 records issued share capital of "
        "₦100,000,000, held as to 80 per cent by Adegoke Segun Babatunde, 10 per cent by Adegoke Mary Ayoboade "
        "and 10 per cent by Adegoke Denisa. The amount recognised is ₦1,000,000. The unpaid balance of "
        "₦99,000,000 is presented as a deduction from equity, in accordance with IAS 32, and not as a receivable. "
        "No proceeds were received in the year. At 31 December 2022 the issued and fully paid capital was ₦1,000,000."
    ))
    h2(doc, "19. Dividends")
    body(doc, f"Dividends paid during the year were ₦{d['dividends']:,.2f}. They are charged to retained earnings. No further dividend is proposed.")
    h2(doc, "20. Contingent liability")
    body(doc, (
        "No output value added tax has been recognised on residential land and housing sales. If the tax authority "
        "takes a different view, an exposure may arise. It has not been accrued. No adjusting event has occurred "
        "between the reporting date and the date of approval of these financial statements."
    ))


def main():
    root = Path(__file__).resolve().parents[1]
    for year in (2023, 2024, 2025):
        build(year, root / f"BAAY_PROJECTS_LIMITED_AFS_{year}.docx")


if __name__ == "__main__":
    main()
