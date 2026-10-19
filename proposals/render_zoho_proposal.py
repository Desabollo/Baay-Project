"""Render the Zoho Books engagement proposal to Word and an unlocked PDF."""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from reportlab.lib.colors import Color, white
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    HRFlowable,
    KeepTogether,
    ListFlowable,
    ListItem,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Flowable,
)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from write_zoho_proposal import (  # noqa: E402
    CREAM,
    DATE,
    FEE,
    GOLD,
    INK,
    LINE,
    MUTED,
    OUT_DOCX,
    OUT_PDF,
    PALE,
    REF,
    TEAL,
    TEAL_MID,
    TOTAL,
    VALID,
    VAT,
    WHITE,
    blocks,
    naira,
)

PAGE_W = 17.4  # cm, inside 1.8 cm side margins


def rgb(value):
    value = value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def rl(value):
    value = value.lstrip("#")
    return Color(int(value[0:2], 16) / 255, int(value[2:4], 16) / 255, int(value[4:6], 16) / 255)


# ---------------------------------------------------------------------------
# Word
# ---------------------------------------------------------------------------

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
        run.font.color.rgb = rgb(color)
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


def set_cell_margins(cell, top=40, bottom=40, left=70, right=70):
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
    tblPr = table._tbl.tblPr
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
    grid = table._tbl.find(qn("w:tblGrid"))
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


def para_border_bottom(paragraph, color=GOLD, sz="12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_field(paragraph, instruction, size=8, color=MUTED):
    run = paragraph.add_run()
    set_run_font(run, size=size, color=color)
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


def cell_text(cell, text, *, size=9, bold=False, color=INK, align="left", fill=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }[align]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run("" if text is None else str(text))
    set_run_font(run, size=size, bold=bold, color=color)
    set_cell_margins(cell)
    set_valign(cell, "center")
    if fill:
        shade(cell, fill)
    return p


def is_numeric(value):
    text = str(value).replace("₦", "").replace(",", "").replace(" ", "").strip()
    return bool(text) and all(ch.isdigit() or ch == "." for ch in text)


def column_numeric(rows, index):
    values = [row[index] for row in rows if str(row[index]).strip()]
    return bool(values) and all(is_numeric(value) for value in values)


def add_table(doc, headers, rows, fractions, emphasis=None):
    emphasis = set(emphasis or [])
    widths = [Cm(PAGE_W * f) for f in fractions]
    numeric_cols = {i for i in range(1, len(headers)) if column_numeric(rows, i)}
    data = [headers] + rows
    table = doc.add_table(rows=len(data), cols=len(headers))
    table.alignment = 1
    set_table_widths(table, widths)
    set_table_borders(table)
    for r, row in enumerate(data):
        prevent_row_split(table.rows[r], header=(r == 0))
        for c, value in enumerate(row):
            align = "right" if c in numeric_cols else "left"
            if r == 0:
                cell_text(table.cell(r, c), value, size=8, bold=True, color=WHITE, align="right" if c in numeric_cols else "left", fill=TEAL)
            else:
                is_total = (r - 1) in emphasis
                fill = CREAM if is_total else (PALE if r % 2 == 0 else None)
                cell_text(
                    table.cell(r, c), value,
                    size=8.5, bold=is_total, color=TEAL if is_total else INK,
                    align=align, fill=fill,
                )
    if doc.paragraphs:
        doc.paragraphs[-1].paragraph_format.keep_with_next = True
    return table


def body(doc, text, *, size=10.5, bold=False, space_after=8, space_before=0, center=False, color=INK):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.12
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def h1(doc, text):
    p = doc.add_paragraph(text, style="Heading 1")
    para_border_bottom(p, TEAL, "10")
    return p


def h2(doc, text):
    return doc.add_paragraph(text, style="Heading 2")


def bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.clear()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.08
        run = p.add_run(item)
        set_run_font(run, size=10.5, color=INK)


def callout(doc, text):
    table = doc.add_table(rows=1, cols=1)
    set_table_widths(table, [Cm(PAGE_W)])
    set_table_borders(table, GOLD, "8")
    cell = table.cell(0, 0)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.1
    run = p.add_run(text)
    set_run_font(run, size=10.5, color=INK)
    set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
    shade(cell, CREAM)
    body(doc, "", space_after=6)


def configure(doc):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(1.7)
    section.bottom_margin = Cm(1.6)
    section.header_distance = Cm(0.6)
    section.footer_distance = Cm(0.5)
    section.different_first_page_header_footer = True

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = rgb(INK)
    _rfonts(normal.element.get_or_add_rPr())
    normal.paragraph_format.space_after = Pt(8)

    for style_name, size, color, before, after in (
        ("Heading 1", 16, TEAL, 16, 6),
        ("Heading 2", 12, TEAL_MID, 12, 3),
    ):
        style = doc.styles[style_name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = rgb(color)
        style.font.italic = False
        _rfonts(style.element.get_or_add_rPr())
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.line_spacing = 1.05
        pPr = style.element.get_or_add_pPr()
        outline = pPr.find(qn("w:outlineLvl"))
        if outline is None:
            outline = OxmlElement("w:outlineLvl")
            pPr.append(outline)
        outline.set(qn("w:val"), "0" if style_name == "Heading 1" else "1")

    bullet_style = doc.styles["List Bullet"]
    bullet_style.font.name = "Calibri"
    bullet_style.font.size = Pt(10.5)
    _rfonts(bullet_style.element.get_or_add_rPr())

    header = section.header
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(PAGE_W), WD_TAB_ALIGNMENT.RIGHT)
    r = p.add_run("CLOUDERP ACCOUNTING  ·  BAAY PROJECTS LIMITED")
    set_run_font(r, size=8, bold=True, color=TEAL)
    r = p.add_run("\tZoho Books")
    set_run_font(r, size=8, color=MUTED)
    para_border_bottom(p, TEAL, "8")

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.paragraph_format.space_before = Pt(2)
    fp.paragraph_format.tab_stops.add_tab_stop(Cm(PAGE_W), WD_TAB_ALIGNMENT.RIGHT)
    r = fp.add_run(f"Confidential  ·  {REF}")
    set_run_font(r, size=8, color=MUTED)
    fp.add_run("\t")
    lab = fp.add_run("Page ")
    set_run_font(lab, size=8, color=MUTED)
    add_field(fp, " PAGE ")
    mid = fp.add_run(" of ")
    set_run_font(mid, size=8, color=MUTED)
    add_field(fp, " NUMPAGES ")

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
    core.title = "Baay Projects Limited — Zoho Books engagement proposal"
    core.subject = f"Professional fees {naira(FEE)} exclusive of VAT. Zoho Books subscription excluded."
    core.author = "CloudERP Accounting & Business Solutions"
    core.category = "Engagement proposal"
    core.comments = "Editable Word copy. Headings are Word styles. Amounts are table cells."


def cover_word(doc):
    band = doc.add_table(rows=1, cols=2)
    set_table_widths(band, [Cm(12.2), Cm(PAGE_W - 12.2)])
    no_borders(band)
    left, right = band.cell(0, 0), band.cell(0, 1)
    for cell, text, align in (
        (left, "CLOUDERP ACCOUNTING & BUSINESS SOLUTIONS", "left"),
        (right, "CONFIDENTIAL", "right"),
    ):
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if align == "left" else WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        set_run_font(run, size=9, bold=True, color=WHITE)
        set_cell_margins(cell, top=90, bottom=90, left=120, right=120)
        shade(cell, TEAL)

    body(doc, "ENGAGEMENT PROPOSAL", size=11, bold=True, space_before=28, space_after=4, color=GOLD)
    body(doc, "The subscriber has paid.", size=26, bold=True, space_after=0, color=TEAL)
    body(doc, "The books should know what that means.", size=20, bold=True, space_after=6, color=INK)
    rule = doc.add_paragraph()
    rule.paragraph_format.space_before = Pt(2)
    rule.paragraph_format.space_after = Pt(10)
    para_border_bottom(rule, GOLD, "16")
    body(
        doc,
        "A proposal to set up Zoho Books for Baay Projects Limited. Land, housing, four banks, and a ledger that can tell a payment from a sale.",
        size=12, space_after=14, color=MUTED,
    )

    meta = [
        ["Prepared for", "Baay Projects Limited  ·  RC 1526224"],
        ["Registered office", "No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State"],
        ["For the attention of", "Adegoke Segun Babatunde, Managing Director"],
        ["Prepared by", "CloudERP Accounting & Business Solutions"],
        ["Our office", "168 Awolowo Road, S.W. Ikoyi, Lagos"],
        ["Contact", "+234 708 037 1266   ·   desmond.abollo@clouderpaccounting.com"],
        ["Date", DATE],
        ["Reference", REF],
        ["Valid until", VALID],
    ]
    table = doc.add_table(rows=len(meta), cols=2)
    set_table_widths(table, [Cm(5.2), Cm(PAGE_W - 5.2)])
    set_table_borders(table, LINE, "4")
    for i, (label, value) in enumerate(meta):
        prevent_row_split(table.rows[i])
        cell_text(table.cell(i, 0), label, size=9, bold=True, color=TEAL, fill=PALE)
        cell_text(table.cell(i, 1), value, size=9, fill=PALE)

    body(doc, "", space_after=8)
    quote = doc.add_table(rows=1, cols=1)
    set_table_widths(quote, [Cm(PAGE_W)])
    no_borders(quote)
    cell = quote.cell(0, 0)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run("Most of the naira in the bank is not yet a sale. The books have to be brave enough to say so.")
    set_run_font(run, size=12, italic=True, color=TEAL)
    set_cell_margins(cell, top=100, bottom=80, left=140, right=140)
    shade(cell, CREAM)
    body(doc, "", space_after=8)
    fee = doc.add_table(rows=1, cols=1)
    set_table_widths(fee, [Cm(PAGE_W)])
    no_borders(fee)
    cell = fee.cell(0, 0)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(f"Professional fees  {naira(FEE)}    ·    exclusive of VAT")
    set_run_font(run, size=12, bold=True, color=WHITE)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(1)
    p2.paragraph_format.space_after = Pt(0)
    run = p2.add_run("Zoho Books subscription is not included. You pay Zoho, in the Company’s name. We do not add a margin.")
    set_run_font(run, size=9, color=WHITE)
    set_cell_margins(cell, top=90, bottom=90, left=140, right=140)
    shade(cell, TEAL)
    doc.add_page_break()


def letter_word(doc):
    body(doc, DATE, size=10.5, space_after=10, color=MUTED)
    body(doc, "The Board of Directors", size=11, bold=True, space_after=0)
    body(doc, "Baay Projects Limited", size=11, bold=True, space_after=0)
    body(doc, "No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State", size=10.5, space_after=8)
    body(doc, "For the attention of Adegoke Segun Babatunde, Managing Director", size=10.5, space_after=10, color=TEAL)
    body(doc, "Dear Directors,", size=11, space_after=8)
    body(doc, "You asked for a proposal to put the Company’s accounting on a cloud system. We recommend Zoho Books, and we have written this from the records, not from a brochure.")
    body(doc, "Baay does not run like a shop that sells and banks the same naira on the same day. A subscriber pays for a plot at Green City, or a house at Gorge View Court or Pacific Apartment, often months before the plot is allocated or the keys are handed over. That cash is not yet a sale. Some of the money in the banks is co-ownership, which is not a sale at all. Some of it has simply moved from Providus to Sterling. The client register, the commission sheet and the four bank accounts each tell part of the story. None of them is the ledger.")
    body(doc, f"The professional fee for the setup in this proposal is {naira(FEE)}, exclusive of VAT. The Zoho Books subscription is not included. You will pay that to Zoho, in the Company’s name. We do not add a margin to it.")
    body(doc, "If the scope and the fee are acceptable, the acceptance at the end is enough to engage us. We can start within five working days of that signature and of the first invoice being paid.")
    body(doc, "Yours faithfully,", space_before=8, space_after=2)
    body(doc, "Desmond Abollo", size=12, bold=True, space_after=0, color=TEAL)
    body(doc, "Managing Director", size=10.5, space_after=0)
    body(doc, "CloudERP Accounting & Business Solutions", size=10.5, space_after=10)
    h2(doc, "What follows")
    bullets(doc, [
        "The decision, on one page — the fee, the subscription, and what you will have.",
        "How the business actually runs — estates, instalments, and why cash is not the sale.",
        "What we will set up — Zoho Books Premium, and the rule for each kind of naira.",
        "Scope, eight weeks, fees, what we need from you, and the terms.",
        "Acceptance. One director is enough.",
    ])
    doc.add_page_break()


def signatures_word(doc):
    body(doc, "For CloudERP Accounting & Business Solutions", size=10, bold=True, space_before=8, space_after=2, color=TEAL)
    body(doc, "______________________________", size=12, space_after=0)
    body(doc, "Desmond Abollo", size=11, bold=True, space_after=0)
    body(doc, "Managing Director", size=10, space_after=0)
    body(doc, "Date: ________________________", size=10, space_after=10)

    body(doc, "For Baay Projects Limited", size=10, bold=True, space_after=2, color=TEAL)
    table = doc.add_table(rows=1, cols=2)
    set_table_widths(table, [Cm(PAGE_W / 2), Cm(PAGE_W / 2)])
    no_borders(table)
    lines = (
        ["______________________________", "Adegoke Segun Babatunde", "Managing Director", "Date: ________________________"],
        ["______________________________", "Owolabi Charles Oluwatobi", "Director  (optional)", "Date: ________________________"],
    )
    for cell, parts in zip(table.rows[0].cells, lines):
        cell.text = ""
        set_cell_margins(cell, top=40, bottom=40, left=0, right=80)
        for i, text in enumerate(parts):
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text)
            set_run_font(run, size=11 if i == 0 else 10.5, bold=(i == 1), color=INK)
    body(doc, "", space_after=8)
    body(doc, "Amendments, if any", size=10, bold=True, space_after=2, color=TEAL)
    body(doc, "________________________________________________________________", size=11, space_after=0, color=LINE)
    body(doc, "________________________________________________________________", size=11, space_after=0, color=LINE)
    body(doc, "If this line is blank, the proposal is accepted as written.", size=9, space_after=8, color=MUTED)
    close = doc.add_table(rows=1, cols=1)
    set_table_widths(close, [Cm(PAGE_W)])
    no_borders(close)
    cell = close.cell(0, 0)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run("To start, sign above and return this page. The first invoice is ₦967,500, being ₦900,000 plus VAT. We will be at your office within five working days of that payment. The Zoho subscription is a separate payment, from the Company to Zoho.")
    set_run_font(run, size=10, color=WHITE)
    set_cell_margins(cell, top=90, bottom=90, left=120, right=120)
    shade(cell, TEAL)


def build_docx():
    doc = Document()
    configure(doc)
    cover_word(doc)
    letter_word(doc)
    for block in blocks():
        kind = block[0]
        if kind == "h1":
            if block[1].startswith("9."):
                doc.add_page_break()
            h1(doc, block[1])
        elif kind == "h2":
            h2(doc, block[1])
        elif kind == "p":
            body(doc, block[1])
        elif kind == "bullets":
            bullets(doc, block[1])
        elif kind == "note":
            callout(doc, block[1])
        elif kind == "table":
            _, headers, rows, fractions, emphasis = block
            add_table(doc, headers, rows, fractions, emphasis)
            body(doc, "", space_after=4)
    signatures_word(doc)
    doc.save(OUT_DOCX)
    print("wrote", OUT_DOCX)


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))


class ProposalCanvas(pdfcanvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved = []
        self._fields = []

    def add_field(self, kind, page, spec):
        self._fields.append((page, kind, spec))

    def showPage(self):
        self._saved.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        fields = list(self._fields)
        n = len(self._saved)
        for state in self._saved:
            self.__dict__.update(state)
            self._decorate(n)
            for page, kind, spec in fields:
                if page == self._pageNumber:
                    if kind == "text":
                        self.acroForm.textfield(**spec)
                    else:
                        self.acroForm.checkbox(**spec)
            super().showPage()
        super().save()

    def _decorate(self, n):
        self.saveState()
        if self._pageNumber == 1:
            self.setFillColor(rl(TEAL))
            self.rect(0, A4[1] - 16 * mm, A4[0], 16 * mm, stroke=0, fill=1)
            self.setFillColor(white)
            self.setFont("DejaVuBold", 8)
            self.drawString(16 * mm, A4[1] - 9.5 * mm, "CLOUDERP ACCOUNTING & BUSINESS SOLUTIONS")
            self.drawRightString(A4[0] - 16 * mm, A4[1] - 9.5 * mm, "CONFIDENTIAL")
        else:
            self.setFillColor(rl(TEAL))
            self.setFont("DejaVuBold", 7.5)
            self.drawString(16 * mm, A4[1] - 12 * mm, "CLOUDERP  ·  BAAY PROJECTS LIMITED")
            self.setFillColor(rl(MUTED))
            self.setFont("DejaVu", 7.5)
            self.drawRightString(A4[0] - 16 * mm, A4[1] - 12 * mm, "Zoho Books engagement")
            self.setStrokeColor(rl(TEAL))
            self.setLineWidth(0.8)
            self.line(16 * mm, A4[1] - 14.5 * mm, A4[0] - 16 * mm, A4[1] - 14.5 * mm)
        self.setStrokeColor(rl(LINE))
        self.setLineWidth(0.4)
        self.line(16 * mm, 12 * mm, A4[0] - 16 * mm, 12 * mm)
        self.setFillColor(rl(MUTED))
        self.setFont("DejaVu", 7.5)
        self.drawString(16 * mm, 7 * mm, f"Confidential  ·  {REF}")
        self.drawRightString(A4[0] - 16 * mm, 7 * mm, f"Page {self._pageNumber} of {n}")
        self.restoreState()


class TextSlot(Flowable):
    def __init__(self, name, width, height, multiline=False):
        super().__init__()
        self.name = name
        self.slot_width = width
        self.slot_height = height
        self.multiline = multiline

    def wrap(self, aw, ah):
        self.slot_width = min(self.slot_width, aw)
        return self.slot_width, self.slot_height

    def drawOn(self, canv, x, y, _sW=0):
        canv.saveState()
        canv.setStrokeColor(rl(TEAL_MID))
        canv.setFillColor(white)
        canv.setLineWidth(0.6)
        canv.roundRect(x, y, self.slot_width, self.slot_height, 2, stroke=1, fill=1)
        canv.restoreState()
        if hasattr(canv, "add_field"):
            canv.add_field("text", canv._pageNumber, dict(
                name=self.name,
                x=x + 1.5,
                y=y + 1.5,
                width=self.slot_width - 3,
                height=self.slot_height - 3,
                borderWidth=0,
                forceBorder=False,
                fontName="Helvetica",
                fontSize=8,
                fillColor=white,
                textColor=rl(INK),
                fieldFlags="multiline" if self.multiline else "",
            ))


class CheckSlot(Flowable):
    def __init__(self, name):
        super().__init__()
        self.name = name

    def wrap(self, aw, ah):
        return 14, 14

    def drawOn(self, canv, x, y, _sW=0):
        if hasattr(canv, "add_field"):
            canv.add_field("check", canv._pageNumber, dict(
                name=self.name,
                x=x,
                y=y,
                size=12,
                buttonStyle="check",
                checked=False,
                borderWidth=0.8,
                borderColor=rl(TEAL),
                fillColor=white,
            ))


def pdf_styles():
    s = {}
    s["h1"] = ParagraphStyle("h1", fontName="DejaVuBold", fontSize=13, leading=16, textColor=rl(TEAL), spaceBefore=11, spaceAfter=4)
    s["h2"] = ParagraphStyle("h2", fontName="DejaVuBold", fontSize=10.5, leading=13, textColor=rl(TEAL_MID), spaceBefore=8, spaceAfter=3)
    s["body"] = ParagraphStyle("body", fontName="DejaVu", fontSize=9, leading=12.2, textColor=rl(INK), alignment=TA_JUSTIFY, spaceAfter=6)
    s["small"] = ParagraphStyle("small", fontName="DejaVu", fontSize=8, leading=10.6, textColor=rl(INK), alignment=TA_LEFT, spaceAfter=2)
    s["li"] = ParagraphStyle("li", fontName="DejaVu", fontSize=9, leading=12.0, textColor=rl(INK), alignment=TA_LEFT, spaceAfter=2, spaceBefore=0)
    s["th"] = ParagraphStyle("th", fontName="DejaVuBold", fontSize=7.4, leading=9.4, textColor=white, alignment=TA_LEFT)
    s["td"] = ParagraphStyle("td", fontName="DejaVu", fontSize=7.6, leading=9.8, textColor=rl(INK), alignment=TA_LEFT)
    s["tdr"] = ParagraphStyle("tdr", fontName="DejaVu", fontSize=7.6, leading=9.8, textColor=rl(INK), alignment=TA_RIGHT)
    s["tdb"] = ParagraphStyle("tdb", fontName="DejaVuBold", fontSize=7.6, leading=9.8, textColor=rl(TEAL), alignment=TA_LEFT)
    s["tdbr"] = ParagraphStyle("tdbr", fontName="DejaVuBold", fontSize=7.6, leading=9.8, textColor=rl(TEAL), alignment=TA_RIGHT)
    s["cover_kicker"] = ParagraphStyle("ck", fontName="DejaVuBold", fontSize=9, leading=12, textColor=rl(GOLD), spaceBefore=8, spaceAfter=6)
    s["cover_title"] = ParagraphStyle("ct", fontName="DejaVuBold", fontSize=20, leading=24, textColor=rl(TEAL), spaceAfter=1)
    s["cover_sub"] = ParagraphStyle("cs", fontName="DejaVuBold", fontSize=13.5, leading=17, textColor=rl(INK), spaceAfter=6)
    s["cover_deck"] = ParagraphStyle("cd", fontName="DejaVu", fontSize=9.5, leading=13, textColor=rl(MUTED), spaceBefore=4, spaceAfter=8)
    s["quote"] = ParagraphStyle("q", fontName="DejaVu", fontSize=10, leading=13.5, textColor=rl(TEAL), spaceBefore=2, spaceAfter=2)
    s["label"] = ParagraphStyle("lb", fontName="DejaVuBold", fontSize=8, leading=10.5, textColor=rl(TEAL))
    s["value"] = ParagraphStyle("vl", fontName="DejaVu", fontSize=8, leading=10.5, textColor=rl(INK))
    s["white"] = ParagraphStyle("wh", fontName="DejaVuBold", fontSize=9, leading=12, textColor=white)
    s["white_s"] = ParagraphStyle("whs", fontName="DejaVu", fontSize=8, leading=10.5, textColor=white)
    return s


S = pdf_styles()
PDF_W = A4[0] - 32 * mm


def P(text, style="body"):
    return Paragraph(str(text), S[style])


def pdf_table(headers, rows, fractions, emphasis=None):
    emphasis = set(emphasis or [])
    widths = [PDF_W * f for f in fractions]
    numeric_cols = {i for i in range(1, len(headers)) if column_numeric(rows, i)}
    data = []
    head = []
    for i, h in enumerate(headers):
        sty = ParagraphStyle(
            f"th{i}{abs(hash(h)) % 10000}",
            parent=S["th"],
            alignment=TA_RIGHT if i in numeric_cols else TA_LEFT,
        )
        head.append(Paragraph(str(h), sty))
    data.append(head)
    for r, row in enumerate(rows):
        line = []
        strong = r in emphasis
        for c, val in enumerate(row):
            right = c in numeric_cols
            if strong and right:
                sty = S["tdbr"]
            elif strong:
                sty = S["tdb"]
            elif right:
                sty = S["tdr"]
            else:
                sty = S["td"]
            line.append(Paragraph(str(val), sty))
        data.append(line)
    table = Table(data, colWidths=widths, repeatRows=1)
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), rl(TEAL)),
        ("TEXTCOLOR", (0, 0), (-1, 0), white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("GRID", (0, 0), (-1, -1), 0.3, rl(LINE)),
        ("FONTNAME", (0, 0), (-1, 0), "DejaVuBold"),
    ]
    for r, _row in enumerate(rows, start=1):
        if (r - 1) in emphasis:
            cmds.append(("BACKGROUND", (0, r), (-1, r), rl(CREAM)))
        elif r % 2 == 0:
            cmds.append(("BACKGROUND", (0, r), (-1, r), rl(PALE)))
    table.setStyle(TableStyle(cmds))
    return table


def cover_pdf():
    story = [Spacer(1, 8 * mm)]
    story.append(P("ENGAGEMENT PROPOSAL", "cover_kicker"))
    story.append(P("The subscriber has paid.", "cover_title"))
    story.append(P("The books should know what that means.", "cover_sub"))
    story.append(HRFlowable(width="100%", thickness=1.6, color=rl(GOLD), spaceBefore=1, spaceAfter=6))
    story.append(P(
        "A proposal to set up Zoho Books for Baay Projects Limited. Land, housing, four banks, and a ledger that can tell a payment from a sale.",
        "cover_deck",
    ))
    meta = [
        ["Prepared for", "Baay Projects Limited  ·  RC 1526224"],
        ["Registered office", "No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State"],
        ["For the attention of", "Adegoke Segun Babatunde, Managing Director"],
        ["Prepared by", "CloudERP Accounting &amp; Business Solutions"],
        ["Our office", "168 Awolowo Road, S.W. Ikoyi, Lagos"],
        ["Contact", "+234 708 037 1266    ·    desmond.abollo@clouderpaccounting.com"],
        ["Date", DATE],
        ["Reference", REF],
        ["Valid until", VALID],
    ]
    data = [[P(a, "label"), P(b, "value")] for a, b in meta]
    t = Table(data, colWidths=[42 * mm, PDF_W - 42 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), rl(PALE)),
        ("BOX", (0, 0), (-1, -1), 0.4, rl(LINE)),
        ("LINEBELOW", (0, 0), (-1, -2), 0.2, rl(LINE)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]))
    story.append(t)
    story.append(Spacer(1, 36 * mm))
    quote = Table(
        [[P("Most of the naira in the bank is not yet a sale. The books have to be brave enough to say so.", "quote")]],
        colWidths=[PDF_W],
    )
    quote.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), rl(CREAM)),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(quote)
    story.append(Spacer(1, 4 * mm))
    fee = Table(
        [[
            P(f"Professional fees  {naira(FEE)}    ·    exclusive of VAT", "white"),
        ], [
            P("Zoho Books subscription is not included. You pay Zoho, in the Company’s name. We do not add a margin.", "white_s"),
        ]],
        colWidths=[PDF_W],
    )
    fee.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), rl(TEAL)),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (0, 0), 8),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 8),
        ("TOPPADDING", (0, 1), (-1, 1), 1),
    ]))
    story.append(fee)
    story.append(PageBreak())
    return story


def letter_pdf():
    story = [
        P(DATE, "small"),
        Spacer(1, 3 * mm),
        P("<b>The Board of Directors</b>", "body"),
        P("<b>Baay Projects Limited</b>", "body"),
        P("No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State", "body"),
        Spacer(1, 2 * mm),
        P("For the attention of Adegoke Segun Babatunde, Managing Director", "h2"),
        P("Dear Directors,", "body"),
        P("You asked for a proposal to put the Company’s accounting on a cloud system. We recommend Zoho Books, and we have written this from the records, not from a brochure.", "body"),
        P("Baay does not run like a shop that sells and banks the same naira on the same day. A subscriber pays for a plot at Green City, or a house at Gorge View Court or Pacific Apartment, often months before the plot is allocated or the keys are handed over. That cash is not yet a sale. Some of the money in the banks is co-ownership, which is not a sale at all. Some of it has simply moved from Providus to Sterling. The client register, the commission sheet and the four bank accounts each tell part of the story. None of them is the ledger.", "body"),
        P(f"The professional fee for the setup in this proposal is {naira(FEE)}, exclusive of VAT. The Zoho Books subscription is not included. You will pay that to Zoho, in the Company’s name. We do not add a margin to it.", "body"),
        P("If the scope and the fee are acceptable, the acceptance at the end is enough to engage us. We can start within five working days of that signature and of the first invoice being paid.", "body"),
        Spacer(1, 2 * mm),
        P("Yours faithfully,", "body"),
        P("<b>Desmond Abollo</b>", "body"),
        P("Managing Director", "small"),
        P("CloudERP Accounting &amp; Business Solutions", "small"),
        Spacer(1, 3 * mm),
        P("What follows", "h2"),
    ]
    items = [
        "The decision, on one page — the fee, the subscription, and what you will have.",
        "How the business actually runs — estates, instalments, and why cash is not the sale.",
        "What we will set up — Zoho Books Premium, and the rule for each kind of naira.",
        "Scope, eight weeks, fees, what we need from you, and the terms.",
        "Acceptance. One director is enough.",
    ]
    story.append(ListFlowable(
        [ListItem(P(item, "li"), leftIndent=10) for item in items],
        bulletType="bullet", start="•", leftIndent=14, bulletFontName="DejaVu", bulletFontSize=8,
    ))
    return story


def acceptance_pdf():
    story = [Spacer(1, 2 * mm), P("For CloudERP Accounting &amp; Business Solutions", "h2")]
    story.append(P("Desmond Abollo, Managing Director", "small"))
    story.append(Spacer(1, 1 * mm))
    story.append(P("Signature", "label"))
    story.append(TextSlot("cloud_signature", 85 * mm, 16))
    story.append(Spacer(1, 1.5 * mm))
    story.append(P("Date", "label"))
    story.append(TextSlot("cloud_date", 45 * mm, 14))
    story.append(Spacer(1, 4 * mm))
    story.append(P("For Baay Projects Limited", "h2"))
    pair = Table(
        [[
            [P("Adegoke Segun Babatunde", "small"), P("Managing Director", "label"), Spacer(1, 1 * mm), P("Signature", "label"), TextSlot("baay_sig_1", 78 * mm, 16), Spacer(1, 1.2 * mm), P("Date", "label"), TextSlot("baay_date_1", 45 * mm, 14)],
            [P("Owolabi Charles Oluwatobi", "small"), P("Director, optional", "label"), Spacer(1, 1 * mm), P("Signature", "label"), TextSlot("baay_sig_2", 78 * mm, 16), Spacer(1, 1.2 * mm), P("Date", "label"), TextSlot("baay_date_2", 45 * mm, 14)],
        ]],
        colWidths=[PDF_W / 2, PDF_W / 2],
    )
    # Table cells need flowables, not lists. Build properly.
    return story, pair


def acceptance_table():
    left = [
        P("Adegoke Segun Babatunde", "small"),
        P("Managing Director", "label"),
        Spacer(1, 1.2 * mm),
        P("Signature", "label"),
        TextSlot("baay_sig_1", 80 * mm, 16),
        Spacer(1, 1.4 * mm),
        P("Date", "label"),
        TextSlot("baay_date_1", 42 * mm, 14),
    ]
    right = [
        P("Owolabi Charles Oluwatobi", "small"),
        P("Director, optional", "label"),
        Spacer(1, 1.2 * mm),
        P("Signature", "label"),
        TextSlot("baay_sig_2", 80 * mm, 16),
        Spacer(1, 1.4 * mm),
        P("Date", "label"),
        TextSlot("baay_date_2", 42 * mm, 14),
    ]
    # A cell can hold a list of flowables in reportlab.
    table = Table([[left, right]], colWidths=[PDF_W / 2, PDF_W / 2])
    table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (0, 0), 6),
        ("LEFTPADDING", (1, 0), (1, 0), 6),
        ("RIGHTPADDING", (1, 0), (1, 0), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return table


def build_pdf():
    story = []
    story.extend(cover_pdf())
    story.extend(letter_pdf())
    for block in blocks():
        kind = block[0]
        if kind == "h1":
            if block[1].startswith("9."):
                story.append(PageBreak())
            story.append(P(block[1], "h1"))
            story.append(HRFlowable(width="100%", thickness=0.6, color=rl(TEAL), spaceBefore=0, spaceAfter=4))
        elif kind == "h2":
            story.append(P(block[1], "h2"))
        elif kind == "p":
            story.append(P(block[1], "body"))
        elif kind == "note":
            box = Table([[P(block[1], "body")]], colWidths=[PDF_W])
            box.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), rl(CREAM)),
                ("BOX", (0, 0), (-1, -1), 0.8, rl(GOLD)),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]))
            story.append(Spacer(1, 1 * mm))
            story.append(box)
            story.append(Spacer(1, 2 * mm))
        elif kind == "bullets":
            story.append(ListFlowable(
                [ListItem(P(item, "li"), leftIndent=10) for item in block[1]],
                bulletType="bullet", start="•", leftIndent=14,
                bulletFontName="DejaVu", bulletFontSize=8, spaceBefore=1, spaceAfter=3,
            ))
        elif kind == "table":
            _, headers, rows, fractions, emphasis = block
            story.append(KeepTogether([
                Spacer(1, 1 * mm),
                pdf_table(headers, rows, fractions, emphasis),
                Spacer(1, 2.5 * mm),
            ]))
    story.append(Spacer(1, 2 * mm))
    story.append(P("For CloudERP Accounting &amp; Business Solutions", "h2"))
    story.append(P("Desmond Abollo, Managing Director", "small"))
    story.append(Spacer(1, 1 * mm))
    story.append(P("Signature", "label"))
    story.append(TextSlot("cloud_signature", 88 * mm, 16))
    story.append(Spacer(1, 1.4 * mm))
    story.append(P("Date", "label"))
    story.append(TextSlot("cloud_date", 46 * mm, 14))
    story.append(Spacer(1, 3 * mm))
    story.append(P("For Baay Projects Limited", "h2"))
    story.append(acceptance_table())
    story.append(Spacer(1, 4 * mm))
    story.append(P("A signature above is the acceptance. The professional fee is ₦1,800,000, exclusive of VAT. The Zoho Books subscription is not included.", "small"))
    story.append(Spacer(1, 3 * mm))
    story.append(P("Amendments, if any. Leave blank to accept the proposal as written.", "label"))
    story.append(Spacer(1, 1 * mm))
    story.append(TextSlot("amendments", PDF_W, 22, multiline=True))
    story.append(Spacer(1, 8 * mm))
    close = Table(
        [[P("To start, sign above and return this page. The first invoice is ₦967,500, being ₦900,000 plus VAT. We will be at your office within five working days of that payment. The Zoho subscription is a separate payment, from the Company to Zoho.", "white_s")]],
        colWidths=[PDF_W],
    )
    close.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), rl(TEAL)),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(close)

    doc = BaseDocTemplate(
        str(OUT_PDF),
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=20 * mm,
        bottomMargin=16 * mm,
        title="Baay Projects Limited — Zoho Books engagement proposal",
        author="CloudERP Accounting & Business Solutions",
        subject="Editable PDF. Professional fees N1,800,000 exclusive of VAT. Zoho Books subscription excluded.",
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="body", showBoundary=0)
    doc.addPageTemplates([PageTemplate(id="main", frames=[frame])])
    doc.build(story, canvasmaker=ProposalCanvas)
    print("wrote", OUT_PDF)


def main():
    build_docx()
    build_pdf()


if __name__ == "__main__":
    main()
