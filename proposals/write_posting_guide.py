"""Baay Projects Limited — Zoho Books Premium posting guide.

Word and PDF. Screenshots are Zoho's own help screens, embedded so the
accountant is not sent to a marketing page. Sample names in those screens
are Zoho's, not Baay's.
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, RGBColor, Twips
from PIL import Image as PILImage
from reportlab.lib.colors import Color, white
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    CondPageBreak,
    Frame,
    HRFlowable,
    KeepTogether,
    ListFlowable,
    ListItem,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Flowable,
)

ROOT = Path(__file__).resolve().parents[1]
SCREENS = Path(__file__).resolve().parent / "guide-screens"
OUT_DOCX = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Posting_Guide.docx"
OUT_PDF = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Posting_Guide.pdf"

NAVY = "0B1F33"
TEAL = "0E7C66"
TEAL_DARK = "0A5C4C"
INK = "1C2833"
MUTED = "5E6A72"
PALE = "F4F7F6"
LINE = "D5DEDB"
WARN_BG = "FBF4EE"
WARN = "8C4A2F"
WHITE = "FFFFFF"
ROW = "F7FAF9"
GOLD = "8A6A2F"

DATE = "3 November 2026"
REF = "CEA/BPL/ZB/2026-01"
ORG = "Baay Projects Limited"


def rgb(value):
    value = value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def rl(value):
    value = value.lstrip("#")
    return Color(int(value[0:2], 16) / 255, int(value[2:4], 16) / 255, int(value[4:6], 16) / 255)


def naira(amount):
    return f"₦{amount:,.0f}"


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


def set_run_font(run, name="Calibri", size=10.5, bold=False, color=INK, italic=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = rgb(color)
    _rfonts(run._element.get_or_add_rPr(), name)


def shade_cell(cell, fill):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_cell_borders(cell, color="D5DEDB", sz="4", edges=("top", "left", "bottom", "right")):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in edges:
        el = tcBorders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            tcBorders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)


def set_cell_margin(cell, top=60, bottom=60, left=90, right=90):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = tcMar.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            tcMar.append(node)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")


def set_table_widths(table, widths):
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    total = sum(widths)
    tblW.set(qn("w:w"), str(int(total)))
    tblW.set(qn("w:type"), "dxa")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for child in list(grid):
            grid.remove(child)
    else:
        grid = OxmlElement("w:tblGrid")
        tblPr.addnext(grid)
    for w in widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(w)))
        grid.append(gc)
    for row in table.rows:
        for cell, w in zip(row.cells, widths):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:w"), str(int(w)))
            tcW.set(qn("w:type"), "dxa")


def prevent_row_split(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    trPr.append(cant)


def set_repeat_header(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    trPr.append(tblHeader)


def para_border_bottom(paragraph, color, sz="8"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_page_number(paragraph):
    run = paragraph.add_run()
    set_run_font(run, size=8, color=MUTED)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def add_rich(paragraph, text, size=10.5, color=INK, bold=False, italic=False):
    parts = text.split("**")
    for i, part in enumerate(parts):
        if not part:
            continue
        run = paragraph.add_run(part)
        set_run_font(run, size=size, bold=bold or (i % 2 == 1), color=color, italic=italic)


class WordGuide:
    def __init__(self):
        self.doc = Document()
        self.fig = 0
        self._setup()

    def _setup(self):
        section = self.doc.sections[0]
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(1.7)
        section.right_margin = Cm(1.7)
        section.top_margin = Cm(1.55)
        section.bottom_margin = Cm(1.55)
        section.header_distance = Cm(0.55)
        section.footer_distance = Cm(0.45)
        section.different_first_page_header_footer = True

        self.usable = 17.6
        self.dxa = int(self.usable * 567)

        normal = self.doc.styles["Normal"]
        normal.font.name = "Calibri"
        normal.font.size = Pt(10.5)
        normal.font.color.rgb = rgb(INK)
        _rfonts(normal.element.get_or_add_rPr())
        pf = normal.paragraph_format
        pf.space_after = Pt(8)
        pf.space_before = Pt(0)
        pf.line_spacing = 1.08

        for style_name, size, color, before, after in (
            ("Heading 1", 15, TEAL_DARK, 16, 4),
            ("Heading 2", 12, NAVY, 12, 3),
        ):
            style = self.doc.styles[style_name]
            style.font.name = "Calibri"
            style.font.size = Pt(size)
            style.font.bold = True
            style.font.color.rgb = rgb(color)
            style.font.italic = False
            _rfonts(style.element.get_or_add_rPr())
            style.paragraph_format.space_before = Pt(before)
            style.paragraph_format.space_after = Pt(after)
            style.paragraph_format.line_spacing = 1.05
            style.paragraph_format.keep_with_next = True

        header = section.header
        header.is_linked_to_previous = False
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        hp.paragraph_format.space_after = Pt(2)
        r = hp.add_run("BAAY PROJECTS LIMITED")
        set_run_font(r, size=8, bold=True, color=NAVY)
        r = hp.add_run("   ·   Zoho Books posting guide   ·   For the accountant")
        set_run_font(r, size=8, color=MUTED)
        para_border_bottom(hp, TEAL, "10")

        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fp.paragraph_format.space_before = Pt(2)
        r = fp.add_run("Procedure note  ·  " + REF + "  ·  ")
        set_run_font(r, size=8, color=MUTED)
        add_page_number(fp)
        para_border_bottom(fp, "FFFFFF", "0")
        # top border on footer
        pPr = fp._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        top = OxmlElement("w:top")
        top.set(qn("w:val"), "single")
        top.set(qn("w:sz"), "6")
        top.set(qn("w:space"), "1")
        top.set(qn("w:color"), LINE)
        pBdr.append(top)
        pPr.append(pBdr)

        core = self.doc.core_properties
        core.author = "CloudERP Accounting"
        core.title = "Baay Projects Limited — Zoho Books Premium posting guide"
        core.subject = "Daily and month-end posting for the accountant"
        core.category = "Procedure note"
        core.comments = (
            "Screenshots are from Zoho Books help. Names and currency in the "
            "pictures are Zoho's sample, not Baay's."
        )

    def _p(self, text="", size=10.5, color=INK, bold=False, italic=False,
           before=0, after=8, align="left", keep_next=False, keep_together=False):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.line_spacing = 1.08
        p.paragraph_format.keep_together = keep_together
        p.paragraph_format.keep_with_next = keep_next
        if align == "justify":
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        elif align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if text:
            add_rich(p, text, size=size, color=color, bold=bold, italic=italic)
        return p

    def cover(self):
        banner = self.doc.add_table(rows=1, cols=1)
        banner.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_widths(banner, [self.dxa])
        cell = banner.cell(0, 0)
        shade_cell(cell, NAVY)
        set_cell_borders(cell, NAVY, "2")
        set_cell_margin(cell, 160, 170, 180, 180)
        cell.paragraphs[0].clear()
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("PROCEDURE NOTE")
        set_run_font(r, size=9, bold=True, color="8FCBB8")
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_before = Pt(8)
        p2.paragraph_format.space_after = Pt(2)
        r = p2.add_run("Zoho Books Premium")
        set_run_font(r, size=12, color="D5E8E2")
        p3 = cell.add_paragraph()
        p3.paragraph_format.space_before = Pt(0)
        p3.paragraph_format.space_after = Pt(4)
        r = p3.add_run("Posting guide")
        set_run_font(r, size=28, bold=True, color=WHITE)
        p4 = cell.add_paragraph()
        p4.paragraph_format.space_before = Pt(2)
        p4.paragraph_format.space_after = Pt(0)
        r = p4.add_run("What to post, on which screen, to which account.\nDaily receipts, recognition, and the month-end close.")
        set_run_font(r, size=11, color="D5E8E2")

        self._p("", after=6)
        meta = [
            ("Organisation", ORG),
            ("For", "The accountant — Finance and Accounts"),
            ("Also for", "Client Relations, on the receipt and the status only"),
            ("Date", DATE),
            ("Engagement", REF),
            ("Prepared by", "CloudERP Accounting, Lagos"),
            ("Sits beside", "The Zoho Books chart of accounts, 130 accounts"),
        ]
        table = self.doc.add_table(rows=len(meta), cols=2)
        set_table_widths(table, [int(self.dxa * 0.28), int(self.dxa * 0.72)])
        for i, (label, value) in enumerate(meta):
            left, right = table.rows[i].cells
            shade_cell(left, PALE if i % 2 == 0 else WHITE)
            shade_cell(right, PALE if i % 2 == 0 else WHITE)
            set_cell_borders(left, LINE, "4")
            set_cell_borders(right, LINE, "4")
            set_cell_margin(left, 40, 40, 80, 60)
            set_cell_margin(right, 40, 40, 80, 80)
            left.paragraphs[0].clear()
            right.paragraphs[0].clear()
            add_rich(left.paragraphs[0], label, size=9, bold=True, color=TEAL_DARK)
            add_rich(right.paragraphs[0], value, size=10, color=INK)
            prevent_row_split(table.rows[i])

        self._p("", after=8)
        self.note(
            "Start here. A receipt is not a sale. Land is a sale when the plot is allocated, or the register says Bought. Housing is a sale only when handover is recorded. Co-ownership is never a sale. A transfer between the Company’s own banks is never income. A line you cannot name waits outside profit."
        )
        self._p(
            "The screenshots are Zoho’s own help screens, so you can see the button. The names, dates and currency in the pictures are Zoho’s sample. Use the Baay accounts in the steps beside the picture. If a label in your organisation has moved, follow the account in this guide, not the sample name.",
            size=10, color=MUTED, after=6,
        )
        self.doc.add_page_break()

    def contents(self, items):
        self.h1("Contents")
        self._p("Use the section, not the page, if you are reading this on screen. Each section is one job.", after=8)
        table = self.doc.add_table(rows=len(items), cols=2)
        set_table_widths(table, [int(self.dxa * 0.12), int(self.dxa * 0.88)])
        for i, (num, title) in enumerate(items):
            a, b = table.rows[i].cells
            fill = PALE if i % 2 == 0 else WHITE
            shade_cell(a, fill)
            shade_cell(b, fill)
            set_cell_borders(a, "EEF2F1", "2")
            set_cell_borders(b, "EEF2F1", "2")
            set_cell_margin(a, 36, 36, 70, 40)
            set_cell_margin(b, 36, 36, 40, 70)
            a.paragraphs[0].clear()
            b.paragraphs[0].clear()
            add_rich(a.paragraphs[0], num, size=10, bold=True, color=TEAL)
            add_rich(b.paragraphs[0], title, size=10.5, color=INK)
            prevent_row_split(table.rows[i])
        self._p("", after=4)

    def h1(self, text):
        p = self.doc.add_paragraph(style="Heading 1")
        p.clear()
        add_rich(p, text, size=15, bold=True, color=TEAL_DARK)
        para_border_bottom(p, TEAL, "8")
        p.paragraph_format.keep_with_next = True

    def h2(self, text):
        p = self.doc.add_paragraph(style="Heading 2")
        p.clear()
        add_rich(p, text, size=12, bold=True, color=NAVY)
        p.paragraph_format.keep_with_next = True

    def p(self, text):
        self._p(text, align="justify")

    def bullets(self, items):
        for item in items:
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.45)
            p.paragraph_format.first_line_indent = Cm(-0.35)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            run = p.add_run("•  ")
            set_run_font(run, size=10.5, color=TEAL, bold=True)
            add_rich(p, item, size=10.5, color=INK)

    def _box(self, text, fill, border, label, label_color):
        table = self.doc.add_table(rows=1, cols=1)
        set_table_widths(table, [self.dxa])
        cell = table.cell(0, 0)
        shade_cell(cell, fill)
        set_cell_borders(cell, border, "8")
        set_cell_margin(cell, 70, 70, 110, 110)
        cell.paragraphs[0].clear()
        if label:
            add_rich(cell.paragraphs[0], label, size=8, bold=True, color=label_color)
            body = cell.add_paragraph()
            body.paragraph_format.space_before = Pt(2)
            body.paragraph_format.space_after = Pt(0)
        else:
            body = cell.paragraphs[0]
            body.paragraph_format.space_after = Pt(0)
        add_rich(body, text, size=10, color=INK)
        self._p("", after=8)

    def note(self, text):
        self._box(text, PALE, TEAL, "THE RULE", TEAL_DARK)

    def warn(self, text):
        self._box(text, WARN_BG, WARN, "DO NOT", WARN)

    def table(self, headers, rows, widths=None):
        if widths is None:
            widths = [1 / len(headers)] * len(headers)
        dxa = [int(self.dxa * w) for w in widths]
        # absorb rounding into the last column
        dxa[-1] = self.dxa - sum(dxa[:-1])
        table = self.doc.add_table(rows=1 + len(rows), cols=len(headers))
        set_table_widths(table, dxa)
        for j, header in enumerate(headers):
            cell = table.rows[0].cells[j]
            shade_cell(cell, NAVY)
            set_cell_borders(cell, NAVY, "4")
            set_cell_margin(cell, 46, 46, 70, 70)
            cell.paragraphs[0].clear()
            add_rich(cell.paragraphs[0], header, size=8.5, bold=True, color=WHITE)
        set_repeat_header(table.rows[0])
        prevent_row_split(table.rows[0])
        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                cell = table.rows[i + 1].cells[j]
                shade_cell(cell, WHITE if i % 2 else ROW)
                set_cell_borders(cell, LINE, "4")
                set_cell_margin(cell, 42, 42, 70, 70)
                cell.paragraphs[0].clear()
                add_rich(cell.paragraphs[0], value, size=9, color=INK)
            prevent_row_split(table.rows[i + 1])
        self._p("", after=8)

    def journal(self, rows, caption=None):
        if caption:
            self._p(caption, size=9, italic=True, color=MUTED, after=3, keep_next=True)
        headers = ["Account", "Code", "Debit", "Credit"]
        widths = [0.52, 0.14, 0.17, 0.17]
        dxa = [int(self.dxa * w) for w in widths]
        dxa[-1] = self.dxa - sum(dxa[:-1])
        data = []
        deb = cre = 0
        for code, name, d, c in rows:
            data.append([name, code, f"{d:,.0f}" if d else "", f"{c:,.0f}" if c else ""])
            deb += d
            cre += c
        data.append(["", "", f"{deb:,.0f}", f"{cre:,.0f}"])
        table = self.doc.add_table(rows=1 + len(data), cols=4)
        set_table_widths(table, dxa)
        for j, header in enumerate(headers):
            cell = table.rows[0].cells[j]
            shade_cell(cell, TEAL_DARK)
            set_cell_borders(cell, TEAL_DARK, "4")
            set_cell_margin(cell, 40, 40, 70, 70)
            cell.paragraphs[0].clear()
            add_rich(cell.paragraphs[0], header, size=8, bold=True, color=WHITE)
            if j >= 2:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        prevent_row_split(table.rows[0])
        for i, row in enumerate(data):
            last = i == len(data) - 1
            for j, value in enumerate(row):
                cell = table.rows[i + 1].cells[j]
                shade_cell(cell, "E7F3EF" if last else (WHITE if i % 2 else ROW))
                set_cell_borders(cell, LINE, "4")
                set_cell_margin(cell, 36, 36, 70, 70)
                cell.paragraphs[0].clear()
                add_rich(cell.paragraphs[0], value, size=9, bold=last or j == 1, color=INK)
                if j >= 2:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            prevent_row_split(table.rows[i + 1])
        self._p("", after=6)

    def figure(self, key, caption):
        self.fig += 1
        path = SCREENS / key
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run()
        # Fit width; cap height so a tall screen does not dominate.
        with PILImage.open(path) as im:
            w, h = im.size
        max_w = self.usable
        max_h = 11.2
        scale = min(max_w / (w / 96), max_h / (h / 96), 1)
        # PIL dpi is unreliable. Scale from pixels assuming the image should
        # fill the text width unless that makes it too tall.
        disp_w = max_w
        disp_h = max_w * h / w
        if disp_h > max_h:
            disp_h = max_h
            disp_w = max_h * w / h
        run.add_picture(str(path), width=Cm(disp_w))
        cap = self.doc.add_paragraph()
        cap.paragraph_format.space_before = Pt(1)
        cap.paragraph_format.space_after = Pt(8)
        add_rich(
            cap,
            f"Figure {self.fig}.  {caption}  Zoho Books screen. Names, dates and currency in the picture are Zoho’s sample, not Baay’s.",
            size=8, italic=True, color=MUTED,
        )

    def step(self, number, title, body):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        add_rich(p, f"Step {number}.  {title}", size=11, bold=True, color=NAVY)
        self._p(body, align="justify", after=4)

    def save(self):
        self.doc.save(OUT_DOCX)


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))


def pdf_markup(text):
    parts = text.split("**")
    out = []
    for i, part in enumerate(parts):
        esc = (part.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        out.append(f"<b>{esc}</b>" if i % 2 else esc)
    return "".join(out)


class PdfImage(Flowable):
    def __init__(self, path, caption, number):
        super().__init__()
        self.path = str(path)
        self.caption = caption
        self.number = number
        with PILImage.open(path) as im:
            self.px_w, self.px_h = im.size
        self.dw = 0
        self.dh = 0
        self.lines = []

    def _wrap_caption(self, width):
        text = (
            f"Figure {self.number}.  {self.caption}  "
            "Zoho Books screen. Sample data, not Baay."
        )
        words = text.split()
        lines, cur = [], ""
        for word in words:
            trial = (cur + " " + word).strip()
            if pdfmetrics.stringWidth(trial, "DejaVu", 7.5) <= width:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = word
        if cur:
            lines.append(cur)
        return lines

    def wrap(self, aw, ah):
        self.lines = self._wrap_caption(aw)
        cap_h = 3.4 * mm * len(self.lines) + 2.2 * mm
        max_h = min(102 * mm, max(40 * mm, ah - cap_h - 4 * mm))
        scale = min(aw / self.px_w, max_h / self.px_h)
        self.dw = self.px_w * scale
        self.dh = self.px_h * scale
        self.cap_h = cap_h
        self.width = aw
        self.height = self.dh + cap_h
        return aw, self.height

    def draw(self):
        self.canv.drawImage(
            self.path, 0, self.cap_h, width=self.dw, height=self.dh,
            preserveAspectRatio=True, mask="auto",
        )
        self.canv.setFillColor(rl(MUTED))
        self.canv.setFont("DejaVu", 7.5)
        y = self.cap_h - 4.2 * mm
        for line in self.lines:
            self.canv.drawString(0, y, line)
            y -= 3.4 * mm


def draw_header_footer(canvas, doc):
    canvas.saveState()
    if doc.page > 1:
        canvas.setFillColor(rl(NAVY))
        canvas.setFont("DejaVuBold", 7.5)
        canvas.drawString(17 * mm, A4[1] - 11 * mm, "BAAY PROJECTS LIMITED")
        canvas.setFillColor(rl(MUTED))
        canvas.setFont("DejaVu", 7.5)
        canvas.drawRightString(A4[0] - 17 * mm, A4[1] - 11 * mm, "Zoho Books posting guide")
        canvas.setStrokeColor(rl(TEAL))
        canvas.setLineWidth(1)
        canvas.line(17 * mm, A4[1] - 13.2 * mm, A4[0] - 17 * mm, A4[1] - 13.2 * mm)
        canvas.setStrokeColor(rl(LINE))
        canvas.setLineWidth(0.4)
        canvas.line(17 * mm, 12 * mm, A4[0] - 17 * mm, 12 * mm)
        canvas.setFillColor(rl(MUTED))
        canvas.setFont("DejaVu", 7.5)
        canvas.drawString(17 * mm, 7.2 * mm, f"Procedure note  ·  {REF}")
        canvas.drawRightString(A4[0] - 17 * mm, 7.2 * mm, str(doc.page))
    canvas.restoreState()


class PdfGuide:
    def __init__(self):
        self.story = []
        self.fig = 0
        self.styles = self._styles()

    def _styles(self):
        s = {}
        s["h1"] = ParagraphStyle(
            "h1", fontName="DejaVuBold", fontSize=13, leading=16,
            textColor=rl(TEAL_DARK), spaceBefore=12, spaceAfter=4,
        )
        s["h2"] = ParagraphStyle(
            "h2", fontName="DejaVuBold", fontSize=11, leading=14,
            textColor=rl(NAVY), spaceBefore=9, spaceAfter=3,
        )
        s["body"] = ParagraphStyle(
            "body", fontName="DejaVu", fontSize=9, leading=12.2,
            textColor=rl(INK), spaceAfter=6, alignment=TA_JUSTIFY,
        )
        s["bullet"] = ParagraphStyle(
            "bullet", fontName="DejaVu", fontSize=9, leading=12,
            textColor=rl(INK), leftIndent=12, bulletIndent=0, spaceAfter=2,
        )
        s["caption"] = ParagraphStyle(
            "caption", fontName="DejaVu", fontSize=8, leading=10.5,
            textColor=rl(MUTED), spaceBefore=1, spaceAfter=8,
        )
        s["step"] = ParagraphStyle(
            "step", fontName="DejaVuBold", fontSize=10, leading=13,
            textColor=rl(NAVY), spaceBefore=7, spaceAfter=2,
        )
        s["small"] = ParagraphStyle(
            "small", fontName="DejaVu", fontSize=8, leading=10.4,
            textColor=rl(INK), alignment=TA_LEFT,
        )
        s["smallbold"] = ParagraphStyle(
            "smallbold", fontName="DejaVuBold", fontSize=8, leading=10.4,
            textColor=rl(WHITE),
        )
        s["th"] = ParagraphStyle(
            "th", fontName="DejaVuBold", fontSize=7.5, leading=9.4,
            textColor=rl(WHITE),
        )
        s["td"] = ParagraphStyle(
            "td", fontName="DejaVu", fontSize=7.6, leading=9.6,
            textColor=rl(INK),
        )
        s["tdbold"] = ParagraphStyle(
            "tdbold", fontName="DejaVuBold", fontSize=7.6, leading=9.6,
            textColor=rl(INK),
        )
        s["cover_kicker"] = ParagraphStyle(
            "cover_kicker", fontName="DejaVuBold", fontSize=8, leading=11,
            textColor=rl("8FCBB8"),
        )
        s["cover_sub"] = ParagraphStyle(
            "cover_sub", fontName="DejaVu", fontSize=10, leading=13,
            textColor=rl("D5E8E2"), spaceBefore=6,
        )
        s["cover_title"] = ParagraphStyle(
            "cover_title", fontName="DejaVuBold", fontSize=22, leading=26,
            textColor=rl(WHITE), spaceBefore=2, spaceAfter=2,
        )
        s["meta_l"] = ParagraphStyle(
            "meta_l", fontName="DejaVuBold", fontSize=8, leading=10.5,
            textColor=rl(TEAL_DARK),
        )
        s["meta_r"] = ParagraphStyle(
            "meta_r", fontName="DejaVu", fontSize=8.5, leading=11,
            textColor=rl(INK),
        )
        s["box"] = ParagraphStyle(
            "box", fontName="DejaVu", fontSize=8.5, leading=11.4,
            textColor=rl(INK),
        )
        s["boxlabel"] = ParagraphStyle(
            "boxlabel", fontName="DejaVuBold", fontSize=7.5, leading=10,
            textColor=rl(TEAL_DARK), spaceAfter=2,
        )
        return s

    def cover(self):
        banner = Table(
            [[Paragraph("PROCEDURE NOTE", self.styles["cover_kicker"])],
             [Paragraph("Zoho Books Premium", self.styles["cover_sub"])],
             [Paragraph("Posting guide", self.styles["cover_title"])],
             [Paragraph(
                 "What to post, on which screen, to which account.<br/>Daily receipts, recognition, and the month-end close.",
                 self.styles["cover_sub"],
             )]],
            colWidths=[176 * mm],
        )
        banner.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), rl(NAVY)),
            ("LEFTPADDING", (0, 0), (-1, -1), 12),
            ("RIGHTPADDING", (0, 0), (-1, -1), 12),
            ("TOPPADDING", (0, 0), (0, 0), 12),
            ("BOTTOMPADDING", (0, -1), (-1, -1), 14),
            ("TOPPADDING", (0, 1), (-1, -2), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -2), 1),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        self.story.append(Spacer(1, 4 * mm))
        self.story.append(banner)
        self.story.append(Spacer(1, 5 * mm))
        meta = [
            ("Organisation", ORG),
            ("For", "The accountant — Finance and Accounts"),
            ("Also for", "Client Relations, on the receipt and the status only"),
            ("Date", DATE),
            ("Engagement", REF),
            ("Prepared by", "CloudERP Accounting, Lagos"),
            ("Sits beside", "The Zoho Books chart of accounts, 130 accounts"),
        ]
        data = [
            [Paragraph(a, self.styles["meta_l"]), Paragraph(b, self.styles["meta_r"])]
            for a, b in meta
        ]
        t = Table(data, colWidths=[42 * mm, 134 * mm])
        style = [
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("GRID", (0, 0), (-1, -1), 0.3, rl(LINE)),
        ]
        for i in range(len(meta)):
            if i % 2 == 0:
                style.append(("BACKGROUND", (0, i), (-1, i), rl(PALE)))
        t.setStyle(TableStyle(style))
        self.story.append(t)
        self.story.append(Spacer(1, 4 * mm))
        self.note(
            "Start here. A receipt is not a sale. Land is a sale when the plot is allocated, or the register says Bought. Housing is a sale only when handover is recorded. Co-ownership is never a sale. A transfer between the Company’s own banks is never income. A line you cannot name waits outside profit."
        )
        self.p(
            "The screenshots are Zoho’s own help screens, so you can see the button. The names, dates and currency in the pictures are Zoho’s sample. Use the Baay accounts in the steps beside the picture. If a label in your organisation has moved, follow the account in this guide, not the sample name."
        )
        self.story.append(PageBreak())

    def contents(self, items):
        self.h1("Contents")
        self.p("Use the section, not the page, if you are reading this on screen. Each section is one job.")
        data = [
            [Paragraph(n, self.styles["tdbold"]), Paragraph(title, self.styles["td"])]
            for n, title in items
        ]
        t = Table(data, colWidths=[16 * mm, 160 * mm])
        style = [
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LINEBELOW", (0, 0), (-1, -2), 0.2, rl(LINE)),
        ]
        t.setStyle(TableStyle(style))
        self.story.append(t)
        self.story.append(Spacer(1, 3 * mm))

    def h1(self, text):
        self.story.append(CondPageBreak(40 * mm))
        self.story.append(Paragraph(pdf_markup(text), self.styles["h1"]))
        self.story.append(HRFlowable(width="100%", thickness=1, color=rl(TEAL), spaceAfter=4))

    def h2(self, text):
        self.story.append(CondPageBreak(28 * mm))
        self.story.append(Paragraph(pdf_markup(text), self.styles["h2"]))

    def p(self, text):
        self.story.append(Paragraph(pdf_markup(text), self.styles["body"]))

    def bullets(self, items):
        for item in items:
            self.story.append(Paragraph("•  " + pdf_markup(item), self.styles["bullet"]))
        self.story.append(Spacer(1, 2 * mm))

    def _box(self, text, fill, border, label, label_color):
        label_style = ParagraphStyle(
            "lab", parent=self.styles["boxlabel"], textColor=rl(label_color),
        )
        inner = [
            Paragraph(label, label_style),
            Paragraph(pdf_markup(text), self.styles["box"]),
        ]
        t = Table([[inner]], colWidths=[176 * mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), rl(fill)),
            ("BOX", (0, 0), (-1, -1), 1.25, rl(border)),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        self.story.append(Spacer(1, 1 * mm))
        self.story.append(t)
        self.story.append(Spacer(1, 3 * mm))

    def note(self, text):
        self._box(text, PALE, TEAL, "THE RULE", TEAL_DARK)

    def warn(self, text):
        self._box(text, WARN_BG, WARN, "DO NOT", WARN)

    def table(self, headers, rows, widths=None):
        if widths is None:
            widths = [1 / len(headers)] * len(headers)
        cols = [176 * mm * w for w in widths]
        head = [Paragraph(pdf_markup(h), self.styles["th"]) for h in headers]
        data = [head]
        for row in rows:
            data.append([Paragraph(pdf_markup(c), self.styles["td"]) for c in row])
        t = Table(data, colWidths=cols, repeatRows=1)
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), rl(NAVY)),
            ("TEXTCOLOR", (0, 0), (-1, 0), white),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.3, rl(LINE)),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
        for i in range(1, len(data)):
            if i % 2 == 0:
                style.append(("BACKGROUND", (0, i), (-1, i), rl(ROW)))
        t.setStyle(TableStyle(style))
        self.story.append(t)
        self.story.append(Spacer(1, 3 * mm))

    def journal(self, rows, caption=None):
        head = [
            Paragraph(h, self.styles["th"])
            for h in ["Account", "Code", "Debit", "Credit"]
        ]
        data = [head]
        deb = cre = 0
        for code, name, d, c in rows:
            data.append([
                Paragraph(pdf_markup(name), self.styles["td"]),
                Paragraph(code, self.styles["tdbold"]),
                Paragraph(f"{d:,.0f}" if d else "", self.styles["td"]),
                Paragraph(f"{c:,.0f}" if c else "", self.styles["td"]),
            ])
            deb += d
            cre += c
        data.append([
            Paragraph("", self.styles["tdbold"]),
            Paragraph("", self.styles["tdbold"]),
            Paragraph(f"{deb:,.0f}", self.styles["tdbold"]),
            Paragraph(f"{cre:,.0f}", self.styles["tdbold"]),
        ])
        cols = [92 * mm, 22 * mm, 31 * mm, 31 * mm]
        t = Table(data, colWidths=cols)
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), rl(TEAL_DARK)),
            ("BACKGROUND", (0, -1), (-1, -1), rl("E7F3EF")),
            ("GRID", (0, 0), (-1, -1), 0.3, rl(LINE)),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
        for i in range(1, len(data) - 1):
            if i % 2 == 0:
                style.append(("BACKGROUND", (0, i), (-1, i), rl(ROW)))
        t.setStyle(TableStyle(style))
        block = []
        if caption:
            block.append(Paragraph(pdf_markup(caption), self.styles["caption"]))
        block.append(t)
        block.append(Spacer(1, 2.5 * mm))
        self.story.append(KeepTogether(block))

    def figure(self, key, caption):
        self.fig += 1
        path = SCREENS / key
        self.story.append(CondPageBreak(92 * mm))
        self.story.append(PdfImage(path, caption, self.fig))

    def step(self, number, title, body):
        self.story.append(Paragraph(f"Step {number}.  {pdf_markup(title)}", self.styles["step"]))
        self.story.append(Paragraph(pdf_markup(body), self.styles["body"]))

    def save(self):
        doc = BaseDocTemplate(
            str(OUT_PDF),
            pagesize=A4,
            leftMargin=17 * mm,
            rightMargin=17 * mm,
            topMargin=18 * mm,
            bottomMargin=16 * mm,
            title="Baay Projects Limited — Zoho Books Premium posting guide",
            author="CloudERP Accounting",
            subject="Daily and month-end posting for the accountant",
        )
        frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
        doc.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=draw_header_footer)])
        doc.build(self.story)


# ---------------------------------------------------------------------------
# The guide
# ---------------------------------------------------------------------------

CONTENTS = [
    ("1", "The rule"),
    ("2", "The accounts you will touch"),
    ("3", "Estates are projects"),
    ("4", "Set this up once"),
    ("5", "Every receipt, the same day"),
    ("6", "A subscriber instalment"),
    ("7", "Co-ownership"),
    ("8", "A transfer between the Company’s banks"),
    ("9", "A line you cannot yet name"),
    ("10", "When land is allocated"),
    ("11", "When a house is handed over"),
    ("12", "Ancillary charges, construction and other trading"),
    ("13", "Money going out"),
    ("14", "A refund"),
    ("15", "Related parties, dividends and unpaid shares"),
    ("16", "Month-end journals"),
    ("17", "Reconcile, then lock"),
    ("18", "Five screens before you close"),
    ("19", "What never to post"),
    ("20", "A worked week"),
    ("21", "Who does what"),
]


def write(doc):
    doc.cover()
    doc.contents(CONTENTS)

    # 1
    doc.h1("1.  The rule")
    doc.p("A receipt is not a sale. If you remember one sentence from this guide, remember that.")
    doc.p("Land becomes revenue when control of an identified plot has transferred. That is the day the plot is allocated, or the day the register says Bought. Not before. And not above the balance on that subscriber’s contract. Money received above the contract stays a deposit, or it is refunded. It is not extra revenue.")
    doc.p("Housing becomes revenue only when handover is recorded. A house can be fully paid and still not be a sale. Until the handover note is in the file, the money sits in **Unearned Revenue (2020)**. The software will not do this quietly in the background. A status change is a prompt to Finance. Finance posts the recognition. Nobody else does.")
    doc.p("Co-ownership is an investment receipt. It credits **Co-ownership liabilities (2025)**. It never credits sales, and it is not a retainer. A retainer is a subscriber deposit. Co-ownership is not a subscriber.")
    doc.p("A movement between Providus, First Bank, Sterling and GTBank is a transfer. It cancels. It is not income on the way in, and it is not an expense on the way out. If those two lines are posted as income and expense, profit is wrong in both directions and the cash still looks right. That is the error this step exists to stop.")
    doc.p("A line whose narration does not yet say what it was waits outside profit. A receipt waits in **Receipts not yet identified (2045)**. A payment waits in **Payments not yet identified (1180)**. Naming it is a weekly job, before the lock. Forcing it into income to make a month look tidy is not a close. It is a misstatement.")
    doc.note("Zoho will offer to help. On a bank deposit it may suggest Other Income, and mark the line Recognized. That is the word on the button. The suggestion is wrong for a Baay receipt. Do not accept it. Match a transaction you have already posted, or categorise to the holding account. Do not let the feed decide the sale.")

    # 2
    doc.h1("2.  The accounts you will touch")
    doc.p("The chart has 130 accounts. You will not use most of them on a Monday. These are the ones a receipt, a sale, a cost and a close actually hit. Codes are the four-digit codes on the uploaded chart. If a code in this guide is not in Zoho, stop. Do not create a near-duplicate with a slightly different name. Import the chart first. Section 4 says how.")
    doc.table(
        ["Code", "Account", "What it is for"],
        [
            ["1010", "Providus Bank", "One of the four banks. A receipt is posted to the bank that received it."],
            ["1020", "First Bank of Nigeria", "Same rule."],
            ["1030", "Sterling Bank", "Same rule."],
            ["1040", "Guaranty Trust Bank", "Same rule."],
            ["1050", "Petty Cash", "Office float. A reimbursement is not income."],
            ["1055", "Site imprest", "Cash at a site. Spend it to the cost or expense. The return of the float is not income."],
            ["1100", "Accounts Receivable", "Zoho’s customer account. An invoice debits this. You cannot point an invoice at 1120."],
            ["1120", "Trade receivables — land and property", "Presentation of unpaid land, at the statement date. Not the daily invoice account."],
            ["1125", "Trade receivables — housing", "Presentation of unpaid housing, at the statement date."],
            ["1110", "Trade receivables — construction", "Presentation of unpaid construction, at the statement date."],
            ["1130", "Allowance for expected credit losses", "Credit balance. Five per cent of gross trade receivables. The expense is 6080, not this account."],
            ["1150", "Withholding tax credit receivable", "Tax a customer withheld. Do not short the deposit or the sale."],
            ["1175", "Amounts due from related parties", "Directors and related parties. Not subscriber money."],
            ["1180", "Payments not yet identified", "A payment you cannot yet name. Outside profit."],
            ["1205", "Development inventory", "Development cost not yet on a plot or a unit."],
            ["1210", "Land held for resale", "Plot cost. Stays here until the plot is allocated."],
            ["1220", "Construction work in progress", "A house still being built. Not a cost of sale yet."],
            ["1230", "Completed units held for resale", "A finished unit, before handover."],
            ["1510–1530", "Equipment, vehicles, plant", "Capitalise. Do not expense the purchase."],
            ["1590", "Accumulated depreciation", "Credit balance. The expense is 6070."],
            ["2020", "Unearned Revenue", "Subscriber money received before allocation or handover. The retainer account."],
            ["2025", "Co-ownership liabilities", "Investment receipts. Never revenue."],
            ["2045", "Receipts not yet identified", "A receipt you cannot yet name. Outside profit."],
            ["2110–2140", "Tax payables, by levy", "Companies income tax, tertiary education tax, police trust fund, NASENI. Split the credit."],
            ["2150", "VAT Payable", "Only where a supply is actually taxable. Not residential land or housing."],
            ["2160", "Withholding tax payable", "Tax you withheld from a vendor. Not an expense, and not income."],
            ["2510", "Director loan and long-term project funding", "Funding that is not a current account and not a subscriber receipt."],
            ["3010", "Ordinary share capital", "Issued shares, to the extent paid. Not a subscriber receipt."],
            ["3012", "Amount due on issued shares", "Debit in equity. Unpaid calls. Do not invoice this through receivables."],
            ["3030", "Dividends", "Debit in equity. Not an expense. Not director remuneration."],
            ["4020", "Revenue — land and property sales", "The plot, on allocation. Not on the receipt."],
            ["4025", "Revenue — housing", "The unit, on handover. Not on the transfer of money."],
            ["4040", "Revenue — ancillary charges", "Corner piece, documentation, development charges. Their own income."],
            ["4010", "Revenue — civil engineering and construction", "When that performance is complete. Not a plot."],
            ["5020", "Cost of sales — land released", "The plot cost, on the same day as 4020."],
            ["5025", "Cost of sales — housing released", "The unit cost, on the same day as 4025."],
            ["6035", "Selling commissions", "Gross. Do not net a commission against revenue."],
            ["6050", "Bank Fees and Charges", "Bank charges. Not finance costs, and not negative interest."],
            ["6070", "Depreciation Expense", "20 per cent, straight line, a full year in the year of purchase."],
            ["6080", "Bad Debt", "The impairment expense. Credit 1130."],
            ["6085", "Customer refunds", "Only where the receipt was already income and cannot be reversed against that sale."],
            ["6015", "Director remuneration", "Directors. Not staff costs, not a dividend, not a loan."],
        ],
        [0.16, 0.40, 0.44],
    )
    doc.h2("Accounts that stay at nil")
    doc.p("These are on the chart because Zoho requires them, or because they are headers. Do not post a Baay transaction to them.")
    doc.table(
        ["Code", "Account", "Why it stays quiet"],
        [
            ["1000, 2000, 3000, 4000, 5000, 6000", "Headers", "A header is a group. Post to the account under it."],
            ["4099", "Sales", "Zoho’s sample income account. Estate sales go to 4020 or 4025."],
            ["5090", "Cost of Goods Sold", "Zoho’s sample cost account. Released cost goes to 5020 or 5025."],
            ["1290", "Inventory Asset", "Zoho’s sample stock account. Land is 1210. A unit is 1230 or 1220."],
            ["9090", "Uncategorized", "A bank line left here has not been posted. It is not a holding account."],
            ["3098", "Opening Balance Offset", "Must be nil once opening balances agree. Not a plug for a month that will not balance."],
            ["3019", "Owner’s Equity", "For an unincorporated owner. This company has share capital."],
            ["3090", "Drawings", "Same. A dividend is 3030. A director’s loan is 2510 or 1175."],
            ["1090", "Undeposited Funds", "Clear it the same day if a payment lands here. Do not leave subscriber money in it."],
        ],
        [0.28, 0.28, 0.44],
    )

    # 3
    doc.h1("3.  Estates are projects")
    doc.p("An estate is not an account. A new phase is a new project, not a new line on the chart. Every retainer, invoice, payment and cost journal carries the project. That is how the Managing Director sees collections by estate without a second set of books.")
    doc.p("Create them once, under **Time Tracking > Projects > + New**. The project is a tag on the transaction. It is not a time-billing job. Do not sell a plot by timesheet. The sale is the invoice on allocation or handover.")
    doc.table(
        ["Project", "What it holds"],
        [
            ["Green City — Lantaba", "Land, in phases. Allocation is the sale."],
            ["Green City — Idobi", "Land. Own project, so a phase is not lost in Lantaba."],
            ["Green City — Omu-Epe", "Land. Own project."],
            ["Green City Ibadan", "Land at Akinyele-Moniya. Same rule."],
            ["Heritage Garden", "Residential land. Kept separate from Green City."],
            ["Agbowa", "Residential land. Own project."],
            ["Green Hillside", "Own project. Do not fold it into Green City."],
            ["Pacific Court", "Housing. The sale is the handover, not the transfer."],
            ["Pacific Apartment", "Housing. Same rule."],
            ["Gorge View Court, Gbagada", "Housing. Same rule. Receipts are large. That does not make them revenue."],
            ["Baay Foreshore", "Waterfront land and housing. Own collections."],
        ],
        [0.38, 0.62],
    )
    doc.figure(
        "project-new.png",
        "The new-project screen. The name is the estate, not a subscriber. Currency is naira. Do not copy the sample billing method: recurring billing is not how a plot is sold. If Customer Name will not save blank, stop. Do not attach the whole estate to the first subscriber.",
    )
    doc.p("If a new phase opens after go-live, add a project with the same naming. Do not open an account called “Green City Phase 4”. Do not post that phase to a neighbouring project “for now”. A receipt in the wrong project is as hard to unwind as a receipt in the wrong account.")

    # 4
    doc.h1("4.  Set this up once")
    doc.p("Do this before the first live receipt. A receipt posted into an organisation that is still on Zoho’s sample income account will look tidy and be wrong.")
    doc.h2("The organisation")
    doc.bullets([
        "Nigeria edition. Base currency naira. If the organisation is in another currency, stop. Do not post and translate later.",
        "Premium. Projects, inventory and the fixed-asset register are why this is Premium. If **Time Tracking > Projects** is missing, the plan is not Premium. Stop and check the subscription before you invent a workaround.",
        "Reporting basis for the books: accrual. When a journal asks for a reporting method, choose **Accrual and Cash**, unless this guide says otherwise. Do not choose Cash only for a revenue recognition. That hides the sale from the accrual accounts.",
    ])
    doc.h2("The chart")
    doc.p("Import **BAAY_PROJECTS_LIMITED_Zoho_Books_Chart_of_Accounts.csv**. The path is Accountant > Chart of Accounts > Import. Encoding UTF-8. Map name, code, description, type, parent, currency and status. Do not map Account ID. Currency NGN. Status Active. Where Zoho asks how to treat a duplicate name, choose **Overwrite**.")
    doc.p("A few Zoho accounts cannot take a parent on the import: Accounts Receivable, Accounts Payable, Unearned Revenue, VAT Payable, Retained Earnings, and the sample income, cost and stock accounts among them. They are coded. They are imported without a parent. If a row fails because of a parent, clear the parent on that row and import it again. Do not force it. Do not skip the account and recreate it under a new name.")
    doc.h2("The retainer account")
    doc.p("Open Settings and search for **Retainer Invoices**. The liability account must be **Unearned Revenue (2020)**. If it says Sales, or Accounts Receivable, stop. Every instalment you record will hit the wrong place, and the customer statement will look like a credit on a sale that has not happened. Change it before the first receipt. This is the single setting that protects the rule in section 1.")
    doc.h2("Items")
    doc.p("Create these under **Items**, and point each one at its income account. Tax on residential land and housing is none. That is the Company’s existing treatment. This guide is not a tax ruling. If you are not sure a supply is exempt, leave the tax blank and ask before you select a code. A 7.5 per cent code stays available for supplies that are taxable, including other trading where that is the right answer.")
    doc.table(
        ["Item", "Income account", "Tax"],
        [
            ["Land plot", "4020 Revenue — land and property sales", "None"],
            ["Housing unit", "4025 Revenue — housing", "None"],
            ["Ancillary charge", "4040 Revenue — ancillary charges", "None, unless a specific charge has been confirmed as taxable"],
            ["Construction", "4010 Revenue — civil engineering and construction", "Ask before you select a code"],
            ["Consulting", "4030 Revenue — project management and consulting", "Ask before you select a code"],
            ["Other trading", "4130 Other trading income", "Ask before you select a code"],
        ],
        [0.24, 0.48, 0.28],
    )
    doc.warn("Do not use Zoho’s sample item if its account is Sales (4099). Do not put VAT on a land plot or a housing unit.")
    doc.h2("Customers, and the four banks")
    doc.p("Customers come from the client register, not from the spelling on a bank narration. Search before you add. A second customer for the same subscriber, because the bank wrote the name differently, splits the deposit and the later invoice. If the narration does not match a customer, the receipt waits in 2045 until Client Relations names the person. Do not invent a customer called Transfer, or Cash deposit, or the first word of the narration.")
    doc.p("On the customer, keep the status the register already uses: **Paying, Bought, Paid, Handed over, Refund**. Put the plot or unit number in the notes. The status is what tells Finance whether to recognise. Client Relations changes the status. Finance posts.")
    doc.p("The four bank accounts are 1010 Providus, 1020 First Bank, 1030 Sterling and 1040 GTBank. Opening balances agree to the signed statement of financial position at 31 December 2025, including cash of **₦155,904,157.58** across the banks. Do not re-key the old bank lines. There are some twenty-three thousand of them. This guide starts from that opening. It is not a history load.")
    doc.h2("Who may touch what")
    doc.p("Set the roles before go-live, not after the first argument. Section 21 is the card. The short version: Client Relations can open a subscriber, raise a retainer and change a status. They cannot change the chart, publish a journal, categorise a bank line to income, or unlock a month. Finance posts, matches, recognises and locks. The Managing Director sees the five screens in section 18 and approves. The company secretary can read.")

    # 5
    doc.h1("5.  Every receipt, the same day")
    doc.p("The sales list, the commission record and the client register are not the bank. Client Relations knows the subscriber. Finance knows the narration. The bank knows the naira. Zoho is where those three meet, on the day the money arrives, not three weeks later in a workbook.")
    doc.p("Read the narration, find the customer, find the project, then use this table. Do not start from the suggestion Zoho has painted on the bank line.")
    doc.table(
        ["If this is what happened", "Post this", "Sales?"],
        [
            ["Subscriber instalment. Plot still Paying, or house not handed over.", "Retainer invoice, then Record Payment into the bank that received it, then Match. Dr bank, Cr 2020.", "No"],
            ["Plot allocated, or the register says Bought.", "Invoice, item Land plot, account 4020, dated the allocation. Apply the retainer. Journal the cost, Dr 5020 Cr 1210.", "Yes — the contract, not the cash"],
            ["House handed over.", "Invoice, item Housing unit, account 4025, dated the handover. Apply the retainer. Journal the cost, Dr 5025 Cr 1230.", "Yes — the contract, not the cash"],
            ["House fully paid, handover not yet recorded.", "Leave it in 2020. Paid is not handed over.", "No"],
            ["Corner piece, documentation or development charge earned.", "Invoice line, account 4040, on the day it is earned. Not buried in the plot price.", "Yes — that charge only"],
            ["Co-ownership receipt.", "Categorise or journal to 2025. Not a retainer. Not a customer payment.", "No"],
            ["Money moved from one Baay bank to another.", "Transfer. Match both statement lines to it.", "No"],
            ["Receipt you cannot name.", "Categorise to 2045. Same day.", "No"],
            ["Payment you cannot name.", "Categorise to 1180. Same day.", "No"],
            ["Bank charge.", "Categorise to 6050. Not 7010.", "No"],
            ["Commission paid.", "Bill or expense, account 6035. Do not net it against 4020 or 4025.", "No"],
            ["Refund of a deposit still in 2020.", "Refund the retainer payment, from the bank that pays it.", "No"],
            ["Refund of a sale already recognised.", "Credit note against the invoice, then refund the credit if cash goes back.", "Reverses the sale"],
            ["Dividend paid.", "Journal. Dr 3030, Cr bank. Not 6015.", "No"],
            ["Related party repays what they owe the Company.", "Dr bank, Cr 1175. Not 2020.", "No"],
            ["Call paid on issued shares.", "Dr bank, Cr 3012. Do not invoice it.", "No"],
        ],
        [0.32, 0.52, 0.16],
    )
    doc.h2("Before you leave the receipt")
    doc.bullets([
        "The customer is the register customer, not a new spelling of the narration.",
        "The project is the estate. A receipt with no project cannot be found on the estate screen.",
        "The date is the bank date, not the date you happen to be typing.",
        "The credit is 2020, 2025 or 2045. It is not an income account.",
        "The bank line is Matched. It is not still sitting in Uncategorized.",
        "Sales did not move. Open the transaction and look. Do not assume.",
    ])

    # 6
    doc.h1("6.  A subscriber instalment")
    doc.p("This is the Monday transaction. The figures are an example, so you can see the journals balance. Use the register for the real amount, the real subscriber and the real plot. Do not post these figures.")
    doc.p("**Example.** Providus receives ₦1,500,000. The narration is a Green City subscriber at Lantaba. The plot is still Paying. No withholding tax was deducted. No VAT.")
    doc.step(
        "1",
        "Open the retainer from the project",
        "Go to **Time Tracking > Projects**, open **Green City — Lantaba**, then **New Transaction > Create Retainer Invoice**. Starting from the project puts the estate on the document. The other path is **Sales > Retainer Invoices > + New**, and you must then choose the project yourself. Do not save a retainer with no project.",
    )
    doc.figure(
        "project-menu.png",
        "New Transaction, on the project. Zoho’s article has drawn a box around Create Manual Journal. Ignore the box. For a receipt, choose Create Retainer Invoice. Create Invoice is the allocation or the handover. Create Manual Journal is the cost, later. Do not use Log Time.",
    )
    doc.figure(
        "retainer-new.png",
        "The retainer invoice. In Baay’s books: the customer is the subscriber, the project is the estate, the amount is the instalment, and tax is none. There is no income account on this screen. That is the point.",
    )
    doc.p("Fill it as follows.")
    doc.bullets([
        "**Customer.** The subscriber, already on the register. Search. Do not add a duplicate.",
        "**Retainer invoice date.** The bank date.",
        "**Project.** Green City — Lantaba. If you started from the project, this is already filled. Check it.",
        "**Reference.** The plot number.",
        "**Description.** Instalment — plot still Paying. Name the plot. A description that says only “payment” cannot be reviewed.",
        "**Amount.** 1,500,000. The gross instalment, even if you will deal with withholding tax in a moment.",
        "**Tax.** None.",
    ])
    doc.p("Save and Send if the subscriber should have the document. Otherwise save it. **Save as Draft does not record a deposit.** A draft retainer is a piece of paper. The books have not moved.")
    doc.step(
        "2",
        "Record the payment into the bank that received it",
        "Open the retainer and click **Record Payment**. Deposit To is **Providus Bank (1010)**. Amount 1,500,000. Payment date is the bank date, not today if today is later. Payment mode: bank transfer. Reference: the bank narration, and the plot number. Do not deposit it to Undeposited Funds and leave it there. If you do, clear 1090 the same day to the real bank.",
    )
    doc.figure(
        "retainer-payment.png",
        "Record Payment sits on the retainer. In Zoho’s sample the document is still a draft, and the currency is not naira. Ignore both. Record the payment only for money that has arrived, on the bank date, into the Baay bank that received it.",
    )
    doc.p("The books have now moved. Sales has not.")
    doc.journal(
        [
            ("1010", "Providus Bank", 1_500_000, 0),
            ("2020", "Unearned Revenue", 0, 1_500_000),
        ],
        "What Zoho has posted. 4020 did not move.",
    )
    doc.step(
        "3",
        "Match the bank line. Do not categorise it.",
        "Go to **Banking > Providus Bank > Uncategorized Transactions**. Find the ₦1,500,000. Zoho should offer **Match**, against the retainer payment you just recorded. Open it, confirm the customer, the amount and the date, and match. The line leaves the uncategorised list. It does not create a second credit.",
    )
    doc.figure(
        "bank-match.png",
        "Match is the correct click when the retainer payment already exists. Import Statement, on the same screen, is how a missing feed gets in. Add Transaction is for a transfer or a line that has no Zoho document yet — not for this instalment.",
    )
    doc.warn("If the line says Categorize, and the suggestion is Other Income or Recognized, do not click it. That suggestion credits income. A subscriber who is still Paying has not produced income. Create the retainer and the payment first, then come back and Match.")
    doc.figure(
        "bank-categorize.png",
        "The trap. The badge says Other Income, Recognized, and the button offered is Categorize. For a Baay subscriber that click is wrong. Categorize is for a bank charge, a transfer, co-ownership, or a line you cannot yet name. It is not for a plot instalment.",
    )
    doc.h2("If the bank line arrived first")
    doc.p("Create the retainer and record the payment, then Match. Do not categorise the deposit as a **Customer Payment**. A customer payment with no invoice becomes unused credits in Accounts Receivable. The customer statement then shows a credit on a sale, and Unearned Revenue does not move. That is the wrong bucket. A subscriber instalment is a deposit. The account is 2020.")
    doc.h2("If withholding tax was deducted")
    doc.p("The retainer is still the gross instalment. Record Payment is the net that hit the bank. Then publish a journal for the difference, dated the bank date, reporting method Accrual and Cash:")
    doc.journal(
        [
            ("1150", "Withholding tax credit receivable", 75_000, 0),
            ("2020", "Unearned Revenue", 0, 75_000),
        ],
        "Example only: ₦75,000 withheld on the ₦1,500,000 instalment. Use the credit note, not this figure. Attach it.",
    )
    doc.p("Do not short the deposit. Do not reduce 4020. There is no sale yet. The withholding tax credit is an asset, 1150, until it is used or refunded.")

    # 7
    doc.h1("7.  Co-ownership")
    doc.p("Co-ownership is money received from an investor in a scheme. It is not a sale of land or housing, and it is not a subscriber deposit. At 31 December 2025 the Company held co-ownership liabilities of its own. New receipts join that liability. They do not join revenue.")
    doc.warn("Do not raise a retainer for co-ownership. A retainer credits 2020 and will be applied to a sale later by someone who thinks it is a subscriber. Do not record a customer payment. That creates unused credits in receivables.")
    doc.p("On the bank line, **Categorize**. The account is **Co-ownership liabilities (2025)**. The contact is the investor, as a customer or as the name you can reconcile, not as a new subscriber with a plot. The project is the estate, if the panel allows a project. The notes say co-ownership, not a sale. The date is the bank date.")
    doc.p("If the categorise panel will not take 2025 cleanly, do not force it into income. Publish a journal instead, and then match the bank line to that journal.")
    doc.journal(
        [
            ("1040", "Guaranty Trust Bank", 10_000_000, 0),
            ("2025", "Co-ownership liabilities", 0, 10_000_000),
        ],
        "Example. ₦10,000,000 received into GTBank. Not revenue.",
    )
    doc.p("A return of co-ownership money is the reverse: debit 2025, credit the bank that pays it. It is not a dividend, not director remuneration, and not 6085. Match the withdrawal to that payment. Do not categorise it as a vendor bill unless there is a vendor and a bill.")

    # 8
    doc.h1("8.  A transfer between the Company’s banks")
    doc.p("**Example.** ₦1,500,000 leaves Providus on Wednesday and arrives in Sterling the same day, or the next. The narration may say TRF, or own account, or nothing useful. You know it is a transfer because the same amount leaves one Baay bank and enters another.")
    doc.step(
        "1",
        "Post one transfer, not two incomes",
        "Open **Banking > Providus Bank > Add Transaction**. Choose the transfer option — the one that asks for the other bank account. It is not Deposit, not Expense, and not Customer Payment. The other account is **Sterling Bank (1030)**. Amount 1,500,000. Date is the date it left. Reference is the narration on both statements, if you have both.",
    )
    doc.step(
        "2",
        "Match both statement lines to that one transfer",
        "Match the Providus withdrawal. Match the Sterling deposit. Both lines leave Uncategorized. If the feed has not brought the second line yet, import that statement or wait for the feed, and match it when it arrives. Do not categorise the incoming line as income because the other side has not appeared.",
    )
    doc.journal(
        [
            ("1030", "Sterling Bank", 1_500_000, 0),
            ("1010", "Providus Bank", 0, 1_500_000),
        ],
        "What a transfer posts. Profit does not move. 4020 does not move. 6050 does not move.",
    )
    doc.p("A charge taken off the transfer is not part of the transfer. Post the charge on its own, to **6050 Bank Fees and Charges**. A ₦50 charge is not a reason to call the whole ₦1,500,000 a bank charge, and it is not a finance cost. Finance costs, 7010, are interest and similar charges on a facility.")
    doc.p("Petty cash and the site imprest are the same idea. Cash given to the site is a transfer: debit **1055 Site imprest**, credit the bank. When the site spends it, debit the cost or the expense — 1220 if it is a house still being built, 1210 or 1205 if it is land still held, 6020 or whatever the voucher is if it is an overhead — and credit 1055. When the float comes back, debit the bank and credit 1055. The return is not income.")

    # 9
    doc.h1("9.  A line you cannot yet name")
    doc.p("Some narrations do not say what the money was. That is ordinary. It is not a reason to guess, and it is not a reason to leave the line in Uncategorized until someone has time. Uncategorized is not a holding account. A line left there has not been posted, and the bank in Zoho does not agree to the statement.")
    doc.p("The same day:")
    doc.bullets([
        "A deposit you cannot name: **Categorize** to **Receipts not yet identified (2045)**. Notes: the narration, the date, and “unnamed”.",
        "A withdrawal you cannot name: **Categorize** to **Payments not yet identified (1180)**. Same notes.",
        "Do not use 9090 Uncategorized, 3098 Opening Balance Offset, 4099 Sales, or 4150 Sundry receipts.",
    ])
    doc.p("Naming is a weekly job, done by Finance with Client Relations, before the lock. When you know what it was, journal it out of 2045 or 1180 to the right account, dated the day you identify it if the original month is already locked. If the month is still open, you may recategorise the line. Write what it turned out to be. A line older than the week is named or explained. It is not quietly rolled.")
    doc.p("If the feed did not bring the statement, do not type income from a narration on a phone. Import the statement.")
    doc.figure(
        "bank-statements.png",
        "Documents > Bank Statements. Use this only when the feed is not there. Add the file to the bank, then Match or categorise each line under the rules in this guide. Auto-forward is optional. It does not decide the account.",
    )
    doc.p("The button on the bank screen itself is **Import Statement**, top right of the Match screen. Either path is fine. A missing feed is not a missing rule.")

    # 10
    doc.h1("10.  When land is allocated")
    doc.p("The register changes from Paying to Bought, or Client Relations records the allocation and puts the allocation letter on the customer. That is the sale. The cash date is not. A plot allocated this Friday, paid for in March, is revenue on Friday.")
    doc.p("**Example, continued from section 6.** The Lantaba plot is allocated on Friday. Contract price ₦3,000,000. Retainer already held ₦1,500,000. Cost of the plot, already in land inventory, ₦900,000. These figures are an example.")
    doc.step(
        "1",
        "Raise the invoice for the contract, not for the cash",
        "Go to **Sales > Invoices > + New**. Customer: the subscriber. Project: Green City — Lantaba. Date: the allocation date, not the original receipt date. Item: **Land plot**. Confirm the account on the line is **4020**, not Sales. Amount: 3,000,000. Tax: none. Reference: the plot number. Attach the allocation letter. Save and Send, or mark it Open. **A draft invoice does not recognise revenue.**",
    )
    doc.figure(
        "invoice-new.png",
        "The top of a new invoice. Customer is the subscriber. Invoice date is the allocation date, or the handover date, not the bank date. Ignore the sample tax, the currency and the address. In Baay the tax is none.",
    )
    doc.figure(
        "invoice-lines.png",
        "The lines. Choose the item, then look at Tax before you save. It must be none. The sample shows a sales tax and a PayPal option. Ignore both. Do not leave a tax on a land plot or a housing unit. Save and Send recognises the sale. Save as Draft does not.",
    )
    doc.p("Zoho will debit **Accounts Receivable (1100)**. You cannot point the invoice at 1120. Do not fight the screen, and do not raise a journal instead of an invoice to get the presentation account. The customer statement has to show the plot. 1120 is used at the statement date, in section 16, not on a Friday afternoon.")
    doc.journal(
        [
            ("1100", "Accounts Receivable", 3_000_000, 0),
            ("4020", "Revenue — land and property sales", 0, 3_000_000),
        ],
        "The invoice. Revenue is the contract, ₦3,000,000, not the ₦1,500,000 already in the bank.",
    )
    doc.step(
        "2",
        "Apply the retainer. Do not ask the subscriber to pay twice.",
        "On the invoice, where it shows **Retainers Available**, choose **Apply Now**. The other path is to open the retainer and **Apply to Invoice**, or **More > Use Retainers**. Apply 1,500,000. Not more than the unused retainer. Not more than the invoice. The project on the application is the same project.",
    )
    doc.figure(
        "apply-retainer.png",
        "Apply the retainer to the invoice. Type only the unused amount for this plot, and not more than the invoice balance. Baay retainers have no tax, so the sample note about reversing tax does not apply. Then Apply Retainers. Do not leave the deposit in 2020 after the sale.",
    )
    doc.journal(
        [
            ("2020", "Unearned Revenue", 1_500_000, 0),
            ("1100", "Accounts Receivable", 0, 1_500_000),
        ],
        "The retainer leaves deposits and clears part of the invoice. Sales does not move again.",
    )
    doc.p("Accounts Receivable left on this subscriber, for this plot, is ₦1,500,000. That must equal the unpaid balance on the register. If the plot was fully paid, apply the whole price and the receivable for that plot is nil. Do not leave a receivable for money you already hold as a deposit. Do not recognise more than the contract. Any excess still in 2020 is refunded or left as a deposit until someone has decided what it is. It is not additional 4020.")
    doc.step(
        "3",
        "Release the cost the same day",
        "Go to **Accountant > Manual Journals > + New Journal**. Date: the allocation date. Journal number: let Zoho number it, and put the plot number in the reference. Notes, in this form every time: Land allocation | Green City — Lantaba | plot number | subscriber | cost from the plot register. Reporting method: **Accrual and Cash**. Currency: NGN. Contact: the subscriber, on the lines, if the line allows a contact.",
    )
    doc.journal(
        [
            ("5020", "Cost of sales — land and property released", 900_000, 0),
            ("1210", "Land held for resale", 0, 900_000),
        ],
        "The cost. Same day as the invoice. Do not credit 1290 Inventory Asset, and do not debit 5090 Cost of Goods Sold.",
    )
    doc.p("**Publish.** Save as Draft does not post. Debits must equal credits. Zoho will not let you publish a journal that does not balance. If it will not publish, you have not finished.")
    doc.figure(
        "journal.png",
        "A published journal. This is the shape of the cost release, the depreciation, the tax and the co-ownership journal. In the sample the accounts are rent and a bank, and the currency is not naira. Yours must balance, must say Published, and must name the plot in the notes. Draft is not posted.",
    )
    doc.p("If the plot cost is not yet in 1210, put it there first, from 1205 Development inventory or from the vendor bill that bought it, and then release it. Do not recognise the sale and leave the cost in inventory. Do not release cost on a plot that is still Paying. Cost follows revenue. Revenue follows allocation. Cash follows the bank. They are three different dates, and only the last two are allowed to differ.")

    # 11
    doc.h1("11.  When a house is handed over")
    doc.p("The shape is the same as land. The trigger is different. Housing revenue is the handover, not the transfer of money, and not the word Paid on the register. A unit can be fully paid for months and still sit in 2020. That is correct. Do not “catch it up” because the cash arrived last quarter.")
    doc.p("**Separate example.** A unit at Gorge View Court is handed over. Contract ₦45,000,000. Retainers already held ₦30,000,000. Cost sitting in completed units, ₦18,000,000. These figures are an example, not a unit on the register.")
    doc.bullets([
        "Invoice, dated the handover. Item **Housing unit**. Account **4025**. Tax none. Amount 45,000,000, the contract. Project: Gorge View Court, Gbagada. Attach the handover note.",
        "Apply retainers of 30,000,000. Receivable left: 15,000,000, which must equal the register.",
        "Same day, publish the cost journal: debit **5025**, credit **1230**, 18,000,000. Notes: Housing handover | Gorge View Court | unit number | subscriber | cost from the unit register.",
    ])
    doc.journal(
        [
            ("1100", "Accounts Receivable", 45_000_000, 0),
            ("4025", "Revenue — housing", 0, 45_000_000),
            ("2020", "Unearned Revenue", 30_000_000, 0),
            ("1100", "Accounts Receivable", 0, 30_000_000),
            ("5025", "Cost of sales — housing released", 18_000_000, 0),
            ("1230", "Completed units held for resale", 0, 18_000_000),
        ],
        "Three documents, one handover. The invoice, the application of the retainer, and the cost. Shown together so you can see they balance as a set.",
    )
    doc.p("While the house is being built, cost sits in **1220 Construction work in progress**. A payment to a subcontractor for that house is a bill or a categorised withdrawal to 1220, not to 5010 and not to an overhead. When the unit is complete and held, move it from 1220 to **1230**. Release it to 5025 only on handover. A receipt does not move 1220. A bank rule must not be allowed to credit 4025.")
    doc.warn("Do not let Zoho’s suggestion, or a bank rule, recognise housing because the amount is large. Gorge View receipts are large. Size is not a handover.")

    # 12
    doc.h1("12.  Ancillary charges, construction and other trading")
    doc.p("Corner piece, documentation and development charges are real money, and they are easy to bury inside a plot price. They are item **Ancillary charge**, account **4040**, on the day the charge is earned. They are not a silent addition to 4020, and they are not netted against cost of sales. If the charge is invoiced with the plot, it is its own line on that invoice, with its own account. Look at the line before you save. Zoho will use the item’s account. If someone has changed the item, the line will be wrong and the total will still look right.")
    doc.p("Construction and engineering revenue is **4010**, and only when the related performance is complete. The cost, released on that day, is **5010** for materials and subcontracts or **5030** for site labour and equipment hire. Do not put a plot sale in 4010 because the subscriber is a contractor. Do not put a construction invoice in 4020 because it feels like property.")
    doc.p("Project management and consulting is **4030**. Other trading is **4130**, and its cost is **5040**. A trade must not look like a plot. If you cannot tell, from the invoice, whether you sold a plot or a load of materials, the item is wrong.")
    doc.p("Interest earned on a bank balance is **4110 Interest Income**. A bank charge is not negative interest. It is 6050. Interest on a facility is **7010**, and it is not a bank charge either.")

    # 13
    doc.h1("13.  Money going out")
    doc.p("A withdrawal is a bill, an expense, a transfer, a refund, or a line you cannot name. It is not “money out, so profit”. Match it to a document you have already posted wherever you can. Categorise it only when there is no document, and then only to the account in this section.")
    doc.h2("A vendor bill")
    doc.p("**Purchases > Bills > + New.** The vendor is the vendor, searched first, not a new card for every spelling. The date is the bill date. The **Account** on the line is the thing you bought, not a sales account and not whatever Zoho last used.")
    doc.figure(
        "bill-account.png",
        "A vendor bill from Zoho’s help. Ignore the laptop, the tax and the discount. The column that matters is Account. On a Baay bill that column is 6035, 6020, 6040, 1210, 1220 or 1230 — never Sales, and never an income account.",
    )
    doc.table(
        ["What you are paying", "Account on the bill", "Not this"],
        [
            ["Selling commission", "6035 Selling commissions", "Do not net it against 4020 or 4025. Do not bury it in 6040."],
            ["Subcontractor or materials on a house still being built", "1220 Construction work in progress", "Not 5010, until construction revenue is recognised."],
            ["Land still held", "1210 Land held for resale", "Not 5020, until the plot is allocated."],
            ["Development cost not yet on a plot or unit", "1205 Development inventory", "Not an overhead, if it is inventory."],
            ["Office rent", "6020 Rent Expense", "Not 1220, unless it is genuinely site cost in inventory."],
            ["Audit and tax advisory", "6040 Audit and tax advisory fees", "Not 8010. The fee is not the tax."],
            ["Staff salaries", "6010 Salaries and Employee Wages", "Directors are 6015, not 6010."],
            ["Director remuneration", "6015 Director remuneration", "Not a dividend, not 1175 unless it is unpaid and you are crediting the current account."],
            ["Bank charge, no bill", "6050, categorised from the feed", "Not 7010."],
        ],
        [0.34, 0.36, 0.30],
    )
    doc.p("Pay the bill from the bank that pays it: open the bill, record the payment, Deposit From that bank. Then **Match** the withdrawal to the bill payment. If Zoho offers Categorize and suggests Cost of Goods Sold, that suggestion is the same trap as Other Income, facing the other way. Do not take it.")
    doc.p("Withholding tax you deduct from a vendor is not a discount. Pay the vendor the net. Journal the gross-up, or record the bill at gross and the withholding tax as the balance you still owe to the tax authority, in **2160 Withholding tax payable**. Attach the evidence. Do not expense the tax twice, and do not credit income with it.")
    doc.p("A vendor credit note reduces the bill. **Purchases > Vendor Credits**, against that bill. It does not credit 4020. It does not credit 4150.")

    # 14
    doc.h1("14.  A refund")
    doc.p("There are two refunds. They do not share an account. The status Refund on the register is Client Relations’ note. The books follow the document, not the word.")
    doc.h2("The money was still a deposit")
    doc.p("Open **Sales > Payments Received**, find the retainer payment, and **Refund**. Paid through the bank that sends the money. Date is the date it leaves the bank. This debits **2020**. It is not an expense. Do not use 6085. Do not raise a credit note against a sale that was never recognised. Match the withdrawal to the refund.")
    doc.h2("The sale was already recognised")
    doc.p("Go to **Sales > Credit Notes > + New**. Customer: the subscriber. Invoice: the allocation or handover invoice, not a blank credit. Date: the date of the refund decision, or the date cash leaves, consistently, and say which in the notes. The item and the account follow the original sale — 4020 or 4025. Tax: none. Amount: the amount of the sale you are reversing, not a round figure. **Save as Open.** A draft credit note does not reverse the sale.")
    doc.figure(
        "credit-note.png",
        "The credit note. Link it to the invoice. In Baay’s books the account follows the original item, 4020 or 4025, and tax is none. Then refund the credit from the bank if the cash is going back.",
    )
    doc.p("If cash goes back, refund the credit from the credit note, through the paying bank, and match the withdrawal. If the subscriber is staying and the credit will sit against a later invoice, apply it there. Do not also expense it.")
    doc.p("**6085 Customer refunds** is only where the receipt was already income and cannot be reversed against that sale. That should be rare. If you are about to use 6085, stop and check whether a credit note will do it. A deposit refund is never 6085. A commission clawback is not 6085 either. A clawback reduces 6035, or credits it, against the vendor.")

    # 15
    doc.h1("15.  Related parties, dividends and unpaid shares")
    doc.p("Directors and related parties have their own current account. Their money is not subscriber money. A receipt from a director does not go to 2020 unless that director is, separately, a subscriber paying for a plot — and then only the plot instalment goes to 2020, and the rest goes to the current account. Split the line. Do not choose the larger story and post it all there.")
    doc.table(
        ["What happened", "Debit", "Credit"],
        [
            ["The Company pays a related party, increasing what they owe", "1175 Amounts due from related parties", "The bank"],
            ["A related party repays the Company", "The bank", "1175"],
            ["Long-term director or project funding, not a current account", "The bank, when drawn", "2510 Director loan and long-term project funding"],
            ["A bank facility drawn, proceeds identified", "The bank", "2520 Bank project facility"],
            ["A repayment above the proceeds you have identified", "1190 Loan repayments in excess of proceeds identified", "The bank"],
            ["Director remuneration, paid", "6015 Director remuneration", "The bank"],
            ["Director remuneration, not yet paid", "6015", "1175, or 2010 if it is a payable rather than a current account"],
            ["Dividend declared and paid", "3030 Dividends", "The bank"],
            ["Call received on issued shares", "The bank", "3012 Amount due on issued shares"],
        ],
        [0.40, 0.30, 0.30],
    )
    doc.p("Do not invoice a shareholder through Accounts Receivable for the unpaid amount on issued shares. **3012** is a debit in equity, not a customer. When the call is paid, credit 3012. Do not credit 3010 again, and do not credit 4020.")
    doc.p("A dividend is not staff cost, not director remuneration, and not Drawings. Close 3030 to retained earnings at year-end. Do not journal the year’s profit into **Retained Earnings (3020)** by hand. Zoho closes profit there. A second journal doubles it.")
    doc.warn("Do not use Owner’s Equity (3019) or Drawings (3090). Those accounts are on the chart because Zoho’s sample file has them. This company does not post to them.")

    # 16
    doc.h1("16.  Month-end journals")
    doc.p("Do these after every bank line is matched or parked, and before the lock. Each one is a published journal. Reporting method Accrual and Cash. Currency NGN. Notes say what the figure is and where it came from. An example below is an example. Use the register, the asset register, or the tax computation. Do not use the number in this guide.")
    doc.h2("Depreciation")
    doc.p("The policy already in use is 20 per cent a year, straight line, a full year in the year of purchase. If the fixed-asset register is on, and the assets sit in 1510, 1520 and 1530 with accumulated depreciation in 1590, run depreciation from the asset. Do not also journal it. Two depreciations is not caution. It is a double charge.")
    doc.p("If the register is not on yet, journal it. Debit **6070 Depreciation Expense**. Credit **1590 Accumulated depreciation**. Do not credit the cost account. A ₦2,000,000 asset bought this year is ₦400,000 of depreciation for the year, under that policy. A monthly close takes a twelfth, unless you have agreed to post the year in the last month. Say which you did, in the notes, and do not do both.")
    doc.h2("Expected credit losses")
    doc.p("The Company provides five per cent of gross trade receivables. Take the open invoices from the receivables ageing. Do not include retainer balances. A deposit is a liability, not a debt. Do not provide against 2020.")
    doc.p("Five per cent of that gross is the credit balance **1130** should show. Journal only the movement. If the ageing of open invoices is ₦20,000,000, the allowance required is ₦1,000,000. If 1130 already shows a credit of ₦800,000, the journal is ₦200,000. Use your ageing. This arithmetic is only the shape.")
    doc.journal(
        [
            ("6080", "Bad Debt", 200_000, 0),
            ("1130", "Allowance for expected credit losses", 0, 200_000),
        ],
        "The movement, not the full five per cent again. Reverse the entry if the allowance is already too high.",
    )
    doc.p("A named write-off is a separate decision, against the invoice, and it is not how you clear a deposit. Do not use 1130 to hide a balance you have not looked at.")
    doc.h2("Tax")
    doc.p("Post the tax from the tax computation for the period. This guide does not contain a rate for you to apply to profit, and the journal is not an agreed assessment.")
    doc.p("Debit **8010 Income tax expense — current tax**. Credit, split by levy, **2110** companies income tax, **2120** tertiary education tax, **2130** police trust fund levy, **2140** NASENI levy. The credits must add to the debit. Do not post current tax to 2610. Deferred tax, if it is recognised, is **8020** and **1610** or **2610**, and only when it is recognised. Do not open 2610 to make the current-tax journal look complete.")
    doc.p("The audit and tax advisory fee is 6040, when billed. It is not 8010.")
    doc.h2("Presentation of receivables, and only at the statement date")
    doc.p("During the month, leave open invoices in **Accounts Receivable (1100)**. That is where the customer statement is right. Do not reclassify every invoice on a Friday.")
    doc.p("When you are preparing the statement of financial position, Finance moves the open customer balances into the stream accounts, by journal: debit **1120** for unpaid land, **1125** for unpaid housing, **1110** for unpaid construction, and credit 1100. Customer statements still read the customer. Reverse that journal on the first day of the next period, so the new month does not start with the same receivable in two accounts. If you are not preparing statements, do not do this journal. A monthly reclassification that is not reversed will double the asset the moment someone adds 1100 to 1120.")

    # 17
    doc.h1("17.  Reconcile, then lock")
    doc.p("Reconcile each bank, petty cash and the site imprest. The order is fixed. Match and park first. Reconcile second. Lock third. A lock on a bank that still has uncategorised lines freezes the mess.")
    doc.bullets([
        "Banking > the account. Uncategorized Transactions is nil, or every remaining line is already in 2045 or 1180 with a narration and a date.",
        "The balance in Zoho agrees to the statement. The difference you are allowed to leave is nil.",
        "Petty cash and the site imprest agree to a count, not to a memory.",
        "If the feed did not bring the month, Import Statement, or bring the file in through Documents > Bank Statements, and then match. Do not reconcile to a statement you have not entered.",
    ])
    doc.p("Then lock. **Accountant > Transaction Locking.**")
    doc.figure(
        "lock-list.png",
        "Lock Sales, Purchases, Banking and Accountant. Lock All is the close you want once the four have been reviewed. Do not lock Banking and leave Accountant open so that a late journal can still move a locked bank.",
    )
    doc.figure(
        "lock-date.png",
        "The lock date is the last day of the month you have reconciled. The reason is written, not left as the sample sentence. A locked month stays locked.",
    )
    doc.p("Lock date: the last day of the month. Reason: the month, the person, and the statement date it was reconciled to. A correction after the lock is a journal in the open month, with a note that says which month it belongs to. It is not a quiet edit of the old one. If a lock has to be lifted, the Managing Director approves it, and the reason is written on the unlock. Client Relations cannot do this. A lifted lock that is not put back the same day is not a lift. It is an open month.")

    # 18
    doc.h1("18.  Five screens before you close")
    doc.p("Before you tell the Managing Director the month is closed, these five agree. They are the pack. They are not five more reports for their own sake.")
    doc.table(
        ["Screen", "Where", "What “agreed” means"],
        [
            ["1. Cash by bank", "The four bank accounts, plus 1050 and 1055", "Each balance agrees to the statement or the count. Uncategorized is nil."],
            ["2. Collections by estate", "Reports > Payments Received > Customize. Transaction type: Retainer Payment. This month.", "Every retainer payment has a project. The total agrees to what Client Relations says was collected."],
            ["3. Money that is not yet a sale", "Unearned Revenue, 2020", "It moved only because a retainer was applied to an invoice, or refunded. It did not move because a bank rule credited income."],
            ["4. Co-ownership", "2025, on its own", "New receipts are here. None of them are in 4020 or 4025."],
            ["5. Still unnamed", "2045 and 1180", "Each line has an owner and an age. A line older than the week is named or explained."],
        ],
        [0.24, 0.38, 0.38],
    )
    doc.figure(
        "collections-report.png",
        "Payments Received, customized. Choose Retainer Payment, not All, or the screen mixes instalments with invoices and you cannot see collections. Save it as a custom report once it is right, so next month starts here.",
    )
    doc.p("Then look at profit, briefly, and only to confirm the rule held.")
    doc.bullets([
        "4020 equals the plots allocated in the month, at contract price, not the bank credits.",
        "4025 equals the units handed over in the month, at contract price.",
        "5020 and 5025 moved on those same days.",
        "4099 Sales is nil. 9090 Uncategorized is nil. 3098 Opening Balance Offset is nil.",
        "4040 is the ancillary charges, and you can name each one.",
    ])
    doc.p("If 4020 equals the bank, the month is not closed. It has been posted the old way. Unlock nothing yet. Find the retainer that was categorised as income, reverse it, and post section 6.")

    # 19
    doc.h1("19.  What never to post")
    doc.p("Pin this. The rest of the guide is how. This is the list of things that are wrong even when the software allows them.")
    doc.bullets([
        "A subscriber receipt to 4020, 4025, 4099, 4150 or 4110.",
        "A co-ownership receipt to any income account, or to a retainer.",
        "A transfer as income on one side and an expense on the other.",
        "VAT on residential land or housing.",
        "A deposit refund to 6085.",
        "A dividend to 6015, 6010 or 3090.",
        "A commission netted against revenue.",
        "Cost left in 1210 or 1230 after the related sale is recognised.",
        "Revenue recognised with no cost release the same day.",
        "An unidentified line forced into profit so a month will balance.",
        "A draft treated as posted. Draft does not post. Publish, Record Payment and Match do.",
        "Opening Balance Offset used to force a balance.",
        "A parent forced onto an account Zoho locks.",
        "The year’s profit journalled into Retained Earnings by hand.",
        "A customer created from the first word of a bank narration.",
        "A historical reload of the old bank lines, as part of this close.",
        "Housing revenue, because the receipt was large, or because the register says Paid.",
    ])

    # 20
    doc.h1("20.  A worked week")
    doc.p("One subscriber, one plot, three days. The figures are the examples already used. Post them in a test only if you want to see the buttons. Do not post them into the live organisation.")
    doc.h2("Monday — the instalment")
    doc.p("Providus, ₦1,500,000, Green City — Lantaba, plot still Paying. Retainer invoice. Record Payment to 1010. Match the bank line. Sales does not move.")
    doc.journal(
        [
            ("1010", "Providus Bank", 1_500_000, 0),
            ("2020", "Unearned Revenue", 0, 1_500_000),
        ],
    )
    doc.h2("Wednesday — the transfer")
    doc.p("The same ₦1,500,000 leaves Providus and arrives in Sterling. One transfer. Both lines matched. Profit does not move.")
    doc.journal(
        [
            ("1030", "Sterling Bank", 1_500_000, 0),
            ("1010", "Providus Bank", 0, 1_500_000),
        ],
    )
    doc.h2("Friday — the allocation")
    doc.p("The plot is allocated. Contract ₦3,000,000. Cost in 1210, ₦900,000. Invoice to 4020. Apply the retainer of ₦1,500,000. Publish the cost. Receivable left on the subscriber: ₦1,500,000, equal to the register.")
    doc.journal(
        [
            ("1100", "Accounts Receivable", 3_000_000, 0),
            ("4020", "Revenue — land and property sales", 0, 3_000_000),
            ("2020", "Unearned Revenue", 1_500_000, 0),
            ("1100", "Accounts Receivable", 0, 1_500_000),
            ("5020", "Cost of sales — land and property released", 900_000, 0),
            ("1210", "Land held for resale", 0, 900_000),
        ],
        "Invoice, application and cost, shown as one set. They are three documents in Zoho.",
    )
    doc.p("At the end of Friday, 4020 is ₦3,000,000, not ₦1,500,000. 2020 for this subscriber is nil. The bank is unchanged by the allocation. That is the week, working. If your test shows 4020 of ₦1,500,000, you categorised the Monday receipt as income. Reverse it and do Monday again.")
    doc.h2("The same week, three other lines")
    doc.bullets([
        "₦10,000,000 into GTBank from an investor. Credit 2025. Not a retainer. Section 7.",
        "₦250,000 into First Bank, narration useless. Credit 2045 the same day. Name it this week. Section 9.",
        "₦150,000 commission out of Providus. Debit 6035. Match the bill payment. Do not deduct it from 4020. Section 13.",
    ])

    # 21
    doc.h1("21.  Who does what")
    doc.table(
        ["Person", "Does", "Does not"],
        [
            ["Client Relations", "Opens the subscriber, after searching. Raises the retainer. Records the payment when the money is in the bank. Changes the status: Paying, Bought, Paid, Handed over, Refund. Attaches the receipt, the allocation letter or the handover note.", "Categorise a bank line. Publish a journal. Change the chart. Choose an income account. Unlock a month. Create a customer from a narration."],
            ["Finance and Accounts", "Matches the bank the same day. Posts co-ownership, transfers and unnamed lines. Raises the invoice on allocation or handover, applies the retainer, and publishes the cost. Reconciles. Locks. Pulls the five screens.", "Accept Other Income on a subscriber deposit. Recognise housing because it is paid. Use Opening Balance Offset to force a month."],
            ["Managing Director", "Opens the five screens. Approves a lock, and approves any unlock. Decides a handover where the file is unclear.", "Re-key a receipt that Client Relations has already recorded."],
            ["Company secretary", "Reads. Enough to answer a director.", "Post."],
        ],
        [0.22, 0.40, 0.38],
    )
    doc.p("CloudERP configures and supports during the engagement. Admin rights return to the Company at the end of hypercare, unless you ask us to stay. This note is the one your team keeps. The chart is the list of accounts. The register is the list of subscribers. Zoho is where the three are not allowed to disagree.")
    doc.h2("If a button has moved")
    doc.p("Zoho renames a menu from time to time. The account does not move. Search Settings for Retainer Invoices if you cannot find the liability account. Search the left-hand menu for Projects, Banking, Manual Journals and Transaction Locking. The articles this guide was written against are Zoho’s own: retainer invoices, matching and categorising bank transactions, manual journals, bills, credit notes, documents, and transaction locking. If the article and this note disagree on a button, follow the article for the click and follow this note for the account.")
    doc.p("This is a posting guide. It is not a tax ruling, and it is not a reload of the years already closed. Residential land and housing are not charged VAT, which is the treatment the Company already uses. Where a supply might be taxable, ask before you select a code.")

    doc.h1("Where the screens come from")
    doc.p("Every screenshot in this note is a Zoho Books screen, included so the accountant can see the button without leaving the note. The organisation names, currencies and amounts in those pictures are Zoho’s sample data. Follow the caption, not the sample.")
    doc.table(
        ["Figure", "What it shows", "Where Zoho shows it"],
        [
            ["1", "New project", "Projects"],
            ["2", "New Transaction on a project", "Projects"],
            ["3", "New retainer invoice", "Retainer invoices"],
            ["4", "Record Payment on a retainer", "Retainer invoices"],
            ["5", "Match on the bank feed", "Matching transactions"],
            ["6", "Categorize, and the Other Income suggestion", "Matching transactions"],
            ["7", "Bank statements inbox, if the feed is not there", "Documents"],
            ["8", "New invoice, the header", "Invoices"],
            ["9", "Invoice lines, and the tax column", "Invoices"],
            ["10", "Apply a retainer to an invoice", "Retainer invoices"],
            ["11", "A published manual journal", "Manual journals"],
            ["12", "The Account column on a vendor bill", "Bills"],
            ["13", "New credit note", "Credit notes"],
            ["14", "Transaction locking, by module", "Transaction locking"],
            ["15", "The lock date and the reason", "Transaction locking"],
            ["16", "Payments Received, filtered to retainer payments", "Payments received"],
        ],
        [0.12, 0.46, 0.42],
    )
    doc.p("Prepared for the accountant of Baay Projects Limited. " + DATE + ". CloudERP Accounting, Lagos. Engagement " + REF + ".")


def main():
    missing = [name for name in (
        "project-new.png", "project-menu.png", "retainer-new.png",
        "retainer-payment.png", "bank-match.png", "bank-categorize.png",
        "bank-statements.png", "invoice-new.png", "invoice-lines.png",
        "apply-retainer.png", "journal.png", "bill-account.png",
        "credit-note.png", "lock-list.png", "lock-date.png",
        "collections-report.png",
    ) if not (SCREENS / name).exists()]
    if missing:
        raise SystemExit(f"Missing screenshots: {missing}")
    word = WordGuide()
    write(word)
    word.save()
    pdf = PdfGuide()
    write(pdf)
    pdf.save()
    print(f"docx {OUT_DOCX} figures {word.fig}")
    print(f"pdf  {OUT_PDF} figures {pdf.fig}")


if __name__ == "__main__":
    main()
