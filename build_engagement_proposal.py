#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build the Baay Projects Limited engagement proposal (Word .docx).

Services proposed:
  1. Integrated accounting system on Zoho Books  (fixed fee N2,500,000)
  2. Tax compliance & management services        (monthly retainer)

Source facts: BAAY_PROJECTS_LIMITED_Master_Workbook_2023_2025_updated.xlsx
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ----------------------------------------------------------------------------
# Palette & constants
# ----------------------------------------------------------------------------
FOREST = "1B3A2F"      # primary dark green
FOREST_LT = "2F5D4A"
GOLD = "B08A3E"
CREAM = "F6F3EC"
CREAM_2 = "FBFAF6"
GREY = "5C5C5C"
MID_GREY = "8A8A8A"
LINE = "D9D3C6"
WHITE = "FFFFFF"

BODY_FONT = "Calibri"
HEAD_FONT = "Georgia"

PAGE_W = Cm(21.0)
MARGIN_L = Cm(2.3)
MARGIN_R = Cm(2.3)
CONTENT_W_CM = 16.4
CONTENT_W = Cm(CONTENT_W_CM)

# ----------------------------------------------------------------------------
# Low-level XML helpers
# ----------------------------------------------------------------------------

def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), v)
    return e


def shade(element_pr, fill):
    shd = _el("w:shd", **{"w:val": "clear", "w:color": "auto", "w:fill": fill})
    element_pr.append(shd)


def cell_bg(cell, fill):
    shade(cell._tc.get_or_add_tcPr(), fill)


def cell_margins(cell, top=90, bottom=90, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = _el("w:" + name, **{"w:w": str(val), "w:type": "dxa"})
        mar.append(node)
    tcPr.append(mar)


def cell_borders(cell, color=LINE, sz="4", edges=("top", "bottom", "left", "right")):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    for edge in edges:
        tag = "w:" + edge
        existing = borders.find(qn(tag))
        if existing is not None:
            borders.remove(existing)
        borders.append(_el(tag, **{"w:val": "single", "w:sz": sz, "w:space": "0", "w:color": color}))


def no_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_el("w:" + edge, **{"w:val": "nil"}))
    tblPr.append(borders)


def table_borders(table, color=LINE, sz="4", inside=True):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    edges = ["top", "left", "bottom", "right"]
    if inside:
        edges += ["insideH", "insideV"]
    for edge in edges:
        borders.append(_el("w:" + edge, **{"w:val": "single", "w:sz": sz, "w:space": "0", "w:color": color}))
    tblPr.append(borders)


def set_col_widths(table, widths_cm):
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    for row in table.rows:
        for idx, w in enumerate(widths_cm):
            if idx < len(row.cells):
                row.cells[idx].width = Cm(w)
    # grid
    grid = table._tbl.tblGrid
    for idx, gc in enumerate(grid.findall(qn("w:gridCol"))):
        if idx < len(widths_cm):
            gc.set(qn("w:w"), str(int(Cm(widths_cm[idx]).twips)))


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(_el("w:tblHeader", **{"w:val": "true"}))


def keep_row_together(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(_el("w:cantSplit", **{"w:val": "true"}))


def para_bottom_border(paragraph, color=GOLD, sz="12", space="6"):
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    pbdr.append(_el("w:bottom", **{"w:val": "single", "w:sz": sz, "w:space": space, "w:color": color}))
    pPr.append(pbdr)


def add_field(paragraph, instr, font=BODY_FONT, size=8.5, color=GREY):
    """Insert a Word field (PAGE / NUMPAGES) using well-formed, separate runs."""
    def _run():
        r = paragraph.add_run()
        r.font.name = font
        r.font.size = Pt(size)
        r.font.color.rgb = RGBColor.from_string(color)
        return r

    r1 = _run()
    r1._r.append(_el("w:fldChar", **{"w:fldCharType": "begin"}))
    r2 = _run()
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = f" {instr} "
    r2._r.append(instrText)
    r3 = _run()
    r3._r.append(_el("w:fldChar", **{"w:fldCharType": "separate"}))
    r4 = _run()
    t = OxmlElement("w:t")
    t.text = "1"
    r4._r.append(t)
    r5 = _run()
    r5._r.append(_el("w:fldChar", **{"w:fldCharType": "end"}))
    return r4


# ----------------------------------------------------------------------------
# Text helpers
# ----------------------------------------------------------------------------

def style_run(run, size=10.5, bold=False, italic=False, color=None, font=BODY_FONT,
              spacing=None, caps=False):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if spacing is not None:
        rPr = run._r.get_or_add_rPr()
        rPr.append(_el("w:spacing", **{"w:val": str(int(spacing * 20))}))
    if caps:
        run.font.all_caps = True
    return run


def para(container, text="", size=10.5, bold=False, italic=False, color=GREY, font=BODY_FONT,
         align=None, space_before=0, space_after=6, line=1.28, spacing=None, caps=False,
         indent=None, keep_with_next=False):
    p = container.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if align is not None:
        p.alignment = align
    if indent is not None:
        pf.left_indent = Cm(indent)
    if keep_with_next:
        pf.keep_with_next = True
    if text:
        run = p.add_run(text)
        style_run(run, size=size, bold=bold, italic=italic, color=color, font=font,
                  spacing=spacing, caps=caps)
    return p


def rich(container, parts, size=10.5, color=GREY, font=BODY_FONT, align=None,
         space_before=0, space_after=6, line=1.28, indent=None):
    """parts = list of (text, dict(style overrides))"""
    p = container.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if indent is not None:
        pf.left_indent = Cm(indent)
    if align is not None:
        p.alignment = align
    for text, fmt in parts:
        run = p.add_run(text)
        style_run(run,
                  size=fmt.get("size", size),
                  bold=fmt.get("bold", False),
                  italic=fmt.get("italic", False),
                  color=fmt.get("color", color),
                  font=fmt.get("font", font),
                  spacing=fmt.get("spacing"))
    return p


def h1(doc, number, title, space_before=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(9)
    pf.keep_with_next = True
    if number:
        r = p.add_run(f"{number}   ")
        style_run(r, size=9, bold=True, color=GOLD, font=BODY_FONT, spacing=1.2)
    r = p.add_run(title.upper())
    style_run(r, size=13.5, bold=True, color=FOREST, font=HEAD_FONT, spacing=0.6)
    para_bottom_border(p, color=GOLD, sz="8", space="4")
    return p


def h2(doc, title, space_before=8, color=FOREST):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(3)
    pf.keep_with_next = True
    r = p.add_run(title)
    style_run(r, size=10.5, bold=True, color=color, font=BODY_FONT)
    return p


def label(doc, text, color=GOLD, space_before=0, space_after=3):
    return para(doc, text.upper(), size=8, bold=True, color=color, spacing=1.4,
                space_before=space_before, space_after=space_after, keep_with_next=True)


def bullet(doc, text, bold_lead=None, size=10.5, color=GREY, indent=0.55, space_after=3,
           bullet_char="\u2022", bullet_color=GOLD):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(indent)
    pf.first_line_indent = Cm(-indent * 0.62)
    pf.space_before = Pt(0)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.22
    r = p.add_run(bullet_char + "\t")
    style_run(r, size=size, bold=True, color=bullet_color)
    pf.tab_stops.add_tab_stop(Cm(indent))
    if bold_lead:
        r = p.add_run(bold_lead)
        style_run(r, size=size, bold=True, color=FOREST)
    r = p.add_run(text)
    style_run(r, size=size, color=color)
    return p


def numbered(doc, num, text, bold_lead=None, size=10.5, color=GREY, indent=0.72):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(indent)
    pf.first_line_indent = Cm(-indent * 0.72)
    pf.space_before = Pt(0)
    pf.space_after = Pt(4)
    pf.line_spacing = 1.24
    pf.tab_stops.add_tab_stop(Cm(indent))
    r = p.add_run(f"{num}\t")
    style_run(r, size=size, bold=True, color=GOLD, font=HEAD_FONT)
    if bold_lead:
        r = p.add_run(bold_lead)
        style_run(r, size=size, bold=True, color=FOREST)
    r = p.add_run(text)
    style_run(r, size=size, color=color)
    return p



def numbered_rich(doc, num, parts, size=10.5, color=GREY, indent=0.72, space_after=8):
    """Numbered paragraph with mixed-format runs on the same line."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(indent)
    pf.first_line_indent = Cm(-indent * 0.78)
    pf.space_before = Pt(0)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.26
    pf.tab_stops.add_tab_stop(Cm(indent))
    r = p.add_run(f"{num}\t")
    style_run(r, size=size + 0.5, bold=True, color=GOLD, font=HEAD_FONT)
    for text, fmt in parts:
        r = p.add_run(text)
        style_run(r, size=fmt.get("size", size),
                  bold=fmt.get("bold", False),
                  italic=fmt.get("italic", False),
                  color=fmt.get("color", color))
    return p


def spacer(doc, pts=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1
    r = p.add_run("")
    style_run(r, size=pts)
    return p


def page_break(doc):
    doc.add_page_break()

# ----------------------------------------------------------------------------
# Block builders
# ----------------------------------------------------------------------------

def make_table(doc, rows, cols, widths):
    t = doc.add_table(rows=rows, cols=cols)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(t, widths)
    # explicit total width (dxa) - prevents Word re-flowing columns on open
    tblPr = t._tbl.tblPr
    for existing in tblPr.findall(qn("w:tblW")):
        tblPr.remove(existing)
    total = int(sum(Cm(w).twips for w in widths))
    tblPr.append(_el("w:tblW", **{"w:w": str(total), "w:type": "dxa"}))
    return t


def header_row(table, labels, fill=FOREST, color=WHITE, size=8.6):
    row = table.rows[0]
    repeat_header(row)
    for idx, text in enumerate(labels):
        cell = row.cells[idx]
        cell_bg(cell, fill)
        cell_margins(cell, top=80, bottom=80, left=110, right=110)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(text.upper())
        style_run(r, size=size, bold=True, color=color, spacing=0.7)


def body_cell(cell, text, size=9.3, bold=False, color=GREY, align=None, italic=False,
              space_after=0, font=BODY_FONT, line=1.16):
    cell_margins(cell, top=80, bottom=80, left=110, right=110)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line
    if align is not None:
        p.alignment = align
    if text:
        r = p.add_run(text)
        style_run(r, size=size, bold=bold, color=color, italic=italic, font=font)


def data_table(doc, headers, rows, widths, fills=None, size=9.2, header_size=8.6,
               zebra=True, first_col_bold=False, first_col_color=FOREST):
    t = make_table(doc, len(rows) + 1, len(headers), widths)
    table_borders(t, color=LINE, sz="4")
    header_row(t, headers, size=header_size)
    for r_i, row_vals in enumerate(rows):
        row = t.rows[r_i + 1]
        keep_row_together(row)
        fill = None
        if fills and r_i < len(fills) and fills[r_i]:
            fill = fills[r_i]
        elif zebra and r_i % 2 == 1:
            fill = CREAM_2
        for c_i, val in enumerate(row_vals):
            cell = row.cells[c_i]
            if fill:
                cell_bg(cell, fill)
            body_cell(cell,
                      val,
                      size=size,
                      bold=(first_col_bold and c_i == 0),
                      color=(first_col_color if (first_col_bold and c_i == 0) else GREY))
    return t


def callout(doc, title, body, fill=CREAM, bar=GOLD, width=CONTENT_W_CM, title_color=FOREST,
            body_color=GREY, size=10.0, space_before=6, space_after=8):
    t = make_table(doc, 1, 1, [width])
    no_borders(t)
    cell = t.cell(0, 0)
    cell_bg(cell, fill)
    cell_margins(cell, top=170, bottom=170, left=230, right=230)
    # left accent bar
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    borders.append(_el("w:left", **{"w:val": "single", "w:sz": "24", "w:space": "0", "w:color": bar}))
    borders.append(_el("w:top", **{"w:val": "nil"}))
    borders.append(_el("w:bottom", **{"w:val": "nil"}))
    borders.append(_el("w:right", **{"w:val": "nil"}))
    tcPr.append(borders)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(2)
    p0.paragraph_format.line_spacing = 1.12
    if title:
        r = p0.add_run(title.upper())
        style_run(r, size=7.9, bold=True, color=bar, spacing=1.3)
    p1 = cell.add_paragraph()
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(0)
    p1.paragraph_format.line_spacing = 1.22
    r = p1.add_run(body)
    style_run(r, size=size, color=body_color)
    # spacing after table
    spacer(doc, space_after)
    return t


def stat_cards(doc, cards, cols=4, gap_cm=0.28):
    """cards = list of (value, caption). Borderless grid of shaded cards with thin gaps."""
    card_w = (CONTENT_W_CM - gap_cm * (cols - 1)) / cols
    widths = []
    for i in range(cols):
        widths.append(card_w)
        if i < cols - 1:
            widths.append(gap_cm)
    grid_cols = len(widths)
    rows_n = (len(cards) + cols - 1) // cols
    t = make_table(doc, rows_n, grid_cols, widths)
    no_borders(t)
    for i, (value, caption) in enumerate(cards):
        r_i, c_i = divmod(i, cols)
        cell = t.cell(r_i, c_i * 2)
        cell_bg(cell, CREAM)
        cell_margins(cell, top=150, bottom=150, left=40, right=40)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(value)
        style_run(run, size=14.5, bold=True, color=FOREST, font=HEAD_FONT)
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        p2.paragraph_format.line_spacing = 1.05
        run = p2.add_run(caption)
        style_run(run, size=7.2, bold=True, color=GOLD, caps=True, spacing=0.6)
    spacer(doc, 8)
    return t


def section_banner(doc, kicker, title, subtitle=None):
    """Dark green full-width banner used at the top of key pages."""
    t = make_table(doc, 1, 1, [CONTENT_W_CM])
    no_borders(t)
    cell = t.cell(0, 0)
    cell_bg(cell, FOREST)
    cell_margins(cell, top=260, bottom=260, left=300, right=300)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(kicker.upper())
    style_run(r, size=7.8, bold=True, color="C9A75E", spacing=1.8)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(2 if subtitle else 0)
    p2.paragraph_format.line_spacing = 1.05
    r = p2.add_run(title)
    style_run(r, size=17, bold=True, color=WHITE, font=HEAD_FONT)
    if subtitle:
        p3 = cell.add_paragraph()
        p3.paragraph_format.space_before = Pt(0)
        p3.paragraph_format.space_after = Pt(0)
        p3.paragraph_format.line_spacing = 1.18
        r = p3.add_run(subtitle)
        style_run(r, size=9.4, color="D8E0DA", italic=True)
    spacer(doc, 8)
    return t


def pull_quote(doc, text, attribution=None):
    t = make_table(doc, 1, 1, [CONTENT_W_CM])
    no_borders(t)
    cell = t.cell(0, 0)
    cell_bg(cell, CREAM_2)
    cell_margins(cell, top=200, bottom=200, left=320, right=320)
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    borders.append(_el("w:left", **{"w:val": "single", "w:sz": "24", "w:space": "0", "w:color": GOLD}))
    borders.append(_el("w:top", **{"w:val": "nil"}))
    borders.append(_el("w:bottom", **{"w:val": "nil"}))
    borders.append(_el("w:right", **{"w:val": "nil"}))
    tcPr.append(borders)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.18
    r = p.add_run(text)
    style_run(r, size=12.5, italic=True, color=FOREST, font=HEAD_FONT)
    if attribution:
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_before = Pt(4)
        p2.paragraph_format.space_after = Pt(0)
        r = p2.add_run(attribution.upper())
        style_run(r, size=7.2, bold=True, color=GOLD, spacing=1.2)
    spacer(doc, 8)
    return t


# ----------------------------------------------------------------------------
# Document setup
# ----------------------------------------------------------------------------

doc = Document()

# base styles
normal = doc.styles["Normal"]
normal.font.name = BODY_FONT
normal.font.size = Pt(10.5)
normal.font.color.rgb = RGBColor.from_string(GREY)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.28
_rpr = normal.element.get_or_add_rPr()
_lang = OxmlElement("w:lang")
_lang.set(qn("w:val"), "en-GB")
_lang.set(qn("w:eastAsia"), "en-GB")
_lang.set(qn("w:bidi"), "ar-SA")
_rpr.append(_lang)

section = doc.sections[0]
section.page_width = PAGE_W
section.page_height = Cm(29.7)
section.left_margin = MARGIN_L
section.right_margin = MARGIN_R
section.top_margin = Cm(1.9)
section.bottom_margin = Cm(1.8)
section.header_distance = Cm(0.9)
section.footer_distance = Cm(0.9)
section.different_first_page_header_footer = True

# core properties
cp = doc.core_properties
cp.title = "Engagement Proposal — Zoho Books Implementation & Tax Compliance"
cp.subject = "Baay Projects Limited (RC 1526224)"
cp.author = "Finance & Tax Advisory"
cp.keywords = "Zoho Books, tax compliance, IFRS, Nigeria, Baay Projects"
cp.comments = "Prepared from the Baay Projects Limited 2023-2025 master workbook."

# ----------------------------------------------------------------------------
# Header / footer (non-first pages)
# ----------------------------------------------------------------------------
hdr = section.header
hp = hdr.paragraphs[0]
hp.text = ""
hp.paragraph_format.space_after = Pt(0)
hp.paragraph_format.tab_stops.add_tab_stop(Cm(CONTENT_W_CM), WD_TAB_ALIGNMENT.RIGHT)
r = hp.add_run("BAAY PROJECTS LIMITED")
style_run(r, size=7.6, bold=True, color=FOREST, spacing=1.5)
r = hp.add_run("\tEngagement Proposal  |  October 2026")
style_run(r, size=7.6, color=MID_GREY, spacing=0.9)
para_bottom_border(hp, color=LINE, sz="6", space="2")

ftr = section.footer
fp = ftr.paragraphs[0]
fp.text = ""
fp.paragraph_format.space_before = Pt(0)
fp.paragraph_format.space_after = Pt(0)
fp.paragraph_format.tab_stops.add_tab_stop(Cm(CONTENT_W_CM), WD_TAB_ALIGNMENT.RIGHT)
r = fp.add_run("Confidential — prepared for the Board of Baay Projects Limited")
style_run(r, size=7.4, color=MID_GREY)
r = fp.add_run("\t")
style_run(r, size=7.4, color=MID_GREY)
r = fp.add_run("Page ")
style_run(r, size=7.4, color=MID_GREY)
add_field(fp, " PAGE ")
r = fp.add_run(" of ")
style_run(r, size=7.4, color=MID_GREY)
add_field(fp, " NUMPAGES ")

# ----------------------------------------------------------------------------
# COVER
# ----------------------------------------------------------------------------
# thin gold rule at top
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
para_bottom_border(p, color=GOLD, sz="18", space="1")
r = p.add_run("\u00A0")
style_run(r, size=1)

para(doc, "ENGAGEMENT PROPOSAL", size=9, bold=True, color=GOLD, spacing=3.2,
     space_before=14, space_after=4)
para(doc, "Baay Projects Limited", size=11, bold=True, color=FOREST, spacing=1.0, space_after=2)
para(doc, "RC 1526224   |   Lagos, Nigeria", size=8.8, color=MID_GREY, space_after=26)

para(doc, "From Spreadsheet", size=40, bold=True, color=FOREST, font=HEAD_FONT,
     space_before=0, space_after=0, line=1.0)
para(doc, "to System.", size=40, bold=True, color=GOLD, font=HEAD_FONT,
     space_before=0, space_after=14, line=1.0)

para(doc, "A live, integrated accounting platform on Zoho Books \u2014 and a tax function "
          "that files on time, every time.",
     size=13, color=GREY, font=HEAD_FONT, italic=True, space_before=0, space_after=8, line=1.32)

para(doc, "Prepared for the Board of Directors of Baay Projects Limited: a company that has grown "
          "its revenue 13-fold in three years, and now needs books that can keep up.",
     size=10, color=MID_GREY, space_before=0, space_after=30, line=1.35)

# summary band
t = make_table(doc, 1, 3, [CONTENT_W_CM / 3] * 3)
no_borders(t)
for i, (top, big, sub) in enumerate([
    ("Integrated accounting", "\u20A62.5m", "Zoho Books setup, fixed fee"),
    ("Tax compliance", "From \u20A6250k", "per month, VAT-exclusive"),
    ("Go-live", "8 weeks", "First close, 5 working days"),
]):
    cell = t.cell(0, i)
    cell_bg(cell, FOREST if i == 1 else CREAM)
    cell_margins(cell, top=190, bottom=190, left=170, right=170)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(top.upper())
    style_run(r, size=7.2, bold=True, color=("C9A75E" if i == 1 else GOLD), spacing=1.1)
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(big)
    style_run(r, size=17, bold=True, color=(WHITE if i == 1 else FOREST), font=HEAD_FONT)
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    r = p.add_run(sub)
    style_run(r, size=7.4, color=("D8E0DA" if i == 1 else MID_GREY))


# submitted by block
para(doc, "SUBMITTED BY", size=7.4, bold=True, color=GOLD, spacing=1.8, space_before=165, space_after=3)
para(doc, "[Firm Name]  \u00B7  Chartered Accountants & Tax Advisers", size=9.6, bold=True,
     color=FOREST, space_after=1)
para(doc, "[Engagement Partner]  \u00B7  [email address]  \u00B7  [phone number]", size=8.8,
     color=MID_GREY, space_after=0)
para(doc, "20 October 2026   \u00B7   Proposal valid until 19 November 2026", size=8.8,
     color=MID_GREY, space_before=12, space_after=0)

page_break(doc)

# ----------------------------------------------------------------------------
# 1. LETTER TO THE BOARD
# ----------------------------------------------------------------------------
h1(doc, "01", "A letter to the Board")

para(doc, "Dear Directors,", size=10.5, color=FOREST, space_before=2, space_after=8)

para(doc, "In 2022, Baay Projects Limited turned over \u20A68.5 million. By 2025, the company had "
          "reported \u20A6112.5 million \u2014 a 13-fold increase in three years \u2014 with 219 subscribers "
          "across nine estates, a portfolio contract value of \u20A64.03 billion and \u20A6774.6 million still "
          "owed by clients. Very few Nigerian property companies of your age have grown that fast, that cleanly.",
     space_after=8)

para(doc, "We know this because we have walked every line of your 2023\u20132025 master workbook: all "
          "36 tabs, the four bank accounts, the project register, the IFRS 15 workings, the tax computations. "
          "The work is thorough, and it tells a success story. It also tells an uncomfortable one.",
     space_after=8)

callout(doc, "The uncomfortable truth",
        "Your 2025 receipts had to be reconstructed from Google Drive turnover sheets and realtor commission "
        "registers. One bank account\u2019s statements were rebuilt from the ledgers. The year\u2019s cash book "
        "is, in the workbook\u2019s own words, \u201ca list for testing against bank statements, not a complete "
        "cash book.\u201d That is not a competence problem. It is a tools problem \u2014 and at \u20A6112.5 million "
        "in revenue, it is an expensive one.")

para(doc, "This proposal fixes it in two connected moves.", space_after=6)

rich(doc, [("First, ", {"bold": True, "color": FOREST}),
           ("we put a live, integrated accounting system on Zoho Books under the business, with your "
            "chart of accounts, projects, tax codes and opening balances built from the same workbook "
            "your auditors rely on. Fixed fee: \u20A62,500,000. Go-live in eight weeks.", {})],
     space_after=5)
rich(doc, [("Second, ", {"bold": True, "color": FOREST}),
           ("we take tax off your desk \u2014 VAT on the 21st, WHT on the 21st, PAYE on the 10th, CIT and TET "
            "by 30 June \u2014 on a fixed monthly retainer, with a health check of the 2023\u20132025 filings "
            "in the first 90 days.", {})],
     space_after=10)

para(doc, "One decision, taken now, means your 2027 board pack is produced by the system rather than "
          "assembled by hand \u2014 and no tax deadline is met by guesswork again.", space_after=14)

para(doc, "Yours faithfully,", space_before=4, space_after=18)
para(doc, "[Name]", size=10, bold=True, color=FOREST, space_after=0)
para(doc, "Partner, [Firm Name]", size=8.8, color=MID_GREY, space_after=0)
para(doc, "For and on behalf of the engagement team", size=8.8, color=MID_GREY, space_after=0)

spacer(doc, 12)
p = para(doc, "", space_before=0, space_after=6)
para_bottom_border(p, color=LINE, sz="6", space="1")
rich(doc, [("P.S. ", {"bold": True, "color": GOLD}),
           ("More than 90% of what we need to configure your system is already in the master workbook \u2014 "
            "the chart of accounts, the project register, the tax codes, the opening balances. All that is "
            "missing is the decision to use it.", {"italic": True})],
     size=9.3, space_after=0)

page_break(doc)

# ----------------------------------------------------------------------------
# 2. WHERE YOU ARE
# ----------------------------------------------------------------------------
h1(doc, "02", "The Baay we have studied")

para(doc, "We have not written this proposal from a template. The figures below are taken from your audited "
          "2023\u20132025 financial statements and the master workbook, and they shape every recommendation "
          "that follows.", space_after=10)

stat_cards(doc, [
    ("\u20A6112.5m", "FY2025 revenue"),
    ("13.2\u00D7", "Growth since 2022"),
    ("219", "Subscribers"),
    ("\u20A64.03bn", "Portfolio value"),
], cols=4)

stat_cards(doc, [
    ("\u20A622.6m", "FY2025 PBT"),
    ("\u20A65.4m", "Tax liability"),
    ("4", "Bank accounts"),
    ("\u20A6774.6m", "Client balances"),
], cols=4)

h2(doc, "Three things the numbers are telling you", space_before=6)

numbered_rich(doc, "1", [
    ("Your books are audit-ready, but not decision-ready. ", {"bold": True, "color": FOREST}),
    ("The 36-tab workbook is genuinely good work \u2014 12 mandatory checks, all passing. But it is a "
     "photograph, not a dashboard. By the time it is finished, the decisions it should have informed have "
     "already been made. Gross margin has held at 35.3%, yet nobody can tell you today which of the nine "
     "estates is actually making that margin, plot by plot.", {}),
])

numbered_rich(doc, "2", [
    ("You crossed \u20A6100 million \u2014 and your tax profile changed with it. ", {"bold": True, "color": FOREST}),
    ("2025 turnover of \u20A6112.5 million moved Baay from a 20% medium company to a 30% large company, and "
     "switched on the NASENI levy at 0.25% of profit. The 2025 tax charge was \u20A67.7 million, with "
     "\u20A65.4 million still outstanding at year-end. That is no longer a June filing exercise; it is a "
     "monthly cash-flow discipline.", {}),
])

numbered_rich(doc, "3", [
    ("Delay is quietly taxing your cash. ", {"bold": True, "color": FOREST}),
    ("At least \u20A61.3 million of withholding tax credits from clients is sitting unclaimed against future "
     "CIT. Land sales are VAT-exempt but construction and consultancy are not, and the two are being invoiced "
     "from the same account. Add \u20A6774.6 million of outstanding subscriber balances with no automated "
     "dunning, and the leak is bigger than the professional fee to fix it.", {}),
], space_after=10)

callout(doc, "What we are not proposing",
        "A rip-and-replace ERP. A twelve-month transformation. A new auditor. We are proposing a right-sized "
        "system your accountant can actually run, and a tax function that keeps FIRS and LIRS at arm\u2019s length.")

page_break(doc)

# ----------------------------------------------------------------------------
# 3. SERVICE 1 — ZOHO BOOKS
# ----------------------------------------------------------------------------
section_banner(doc, "03  ·  Service one", "An integrated accounting system on Zoho Books",
               "One source of truth for sales, cash, projects and tax \u2014 built on the same chart of "
               "accounts your auditors already know.")

para(doc, "Baay does not have an accounting problem; it has an accounting \u201csomewhere else\u201d problem. "
          "Sales sit in a client register, receipts in turnover sheets, commissions in a live Google Sheet, "
          "bank statements in four separate portals, and the audited position in a spreadsheet model. Zoho "
          "Books replaces that with one ledger, updated daily, that produces the audit file as a by-product "
          "of doing the work.", space_after=8)

h2(doc, "What we will configure", space_before=4)

data_table(doc,
    ["Area", "What goes in"],
    [
        ["Sales & billing", "Invoice templates, instalment and subscription schedules from the client "
                            "register, automated reminders, receipts and allocation"],
        ["Projects & estates", "Nine estates (Green City phases, Pacific Court, Pacific Apartment, Gorge View, "
                               "Baay Foreshore, Heritage Garden) as trackable projects with budgets"],
        ["Purchases", "Supplier and subcontractor bills, retention tracking, purchase approvals, WHT on vendor "
                      "payments"],
        ["Banking", "Providus, First Bank, Sterling and GTBank \u2014 statements imported or fed, rules-based "
                    "categorisation, daily reconciliation"],
        ["Tax engine", "7.5% VAT (with exempt land supplies separated), WHT at 2% and 5%, PAYE, and the "
                       "schedules that feed the annual CIT/TET return"],
        ["Fixed assets", "IAS 16 register: plant, machinery and vehicles, depreciation to 20%/25% straight line"],
        ["Reports", "Cash position, receivables ageing, project margin, tax liability and a one-page monthly "
                    "board pack"],
        ["Controls", "User roles, approval limits, audit trail, document vault, month-end lock"],
    ],
    widths=[4.2, 12.2], first_col_bold=True, size=8.9)

spacer(doc, 10)

h2(doc, "The implementation: six phases, eight weeks", space_before=0)

data_table(doc,
    ["Phase", "Weeks", "What happens", "Output you sign off"],
    [
        ["1. Mobilise", "Week 1", "Kick-off, access to banks and registers, walkthrough of the master workbook",
         "Diagnostic note and project plan"],
        ["2. Design", "Week 2", "Chart of accounts, tax codes, project and customer structures, approval rules",
         "Accounting manual v1"],
        ["3. Configure", "Weeks 3\u20134", "Zoho Books Nigeria edition built, master data loaded, opening balances at "
         "1 January 2026", "Configuration sign-off"],
        ["4. Catch-up", "Weeks 5\u20136", "2026 transactions posted month by month, banks reconciled, tax schedules "
         "populated", "Reconciled trial balance"],
        ["5. Test & train", "Week 7", "Parallel run against your records, two days of training, SOP handover",
         "UAT sign-off, trained users"],
        ["6. Go live", "Week 8", "Cut-over, first live month-end, 30 days hypercare", "Go-live certificate"],
    ],
    widths=[2.5, 2.1, 7.3, 4.5], first_col_bold=True, size=8.8)

spacer(doc, 8)

callout(doc, "Acceptance criteria \u2014 we are paid on these",
        "The opening trial balance agrees to your audited 31 December 2025 statements to the naira \u00B7 every "
        "bank account reconciles at go-live \u00B7 six users are trained and confident \u00B7 the first month-end "
        "closes in five working days \u00B7 you hold a written procedures manual that survives staff turnover.")

page_break(doc)

# ----------------------------------------------------------------------------
# 4. BEFORE / AFTER
# ----------------------------------------------------------------------------
h1(doc, "04", "What changes on Monday morning")

para(doc, "Systems projects fail when they are sold as software. This one is sold as a change in how the "
          "company runs \u2014 and the difference is easiest to see side by side.", space_after=12)

data_table(doc,
    ["Today", "From go-live"],
    [
        ["Year-end revenue rebuilt from Drive turnover sheets and commission registers",
         "Every invoice raised in-system, with due dates, receipts and reminders"],
        ["Four bank accounts reconciled by hand, months in arrears",
         "Daily bank feeds, rules-based matching, exceptions visible same day"],
        ["VAT and WHT computed after the fact from the accounts",
         "Tax captured at the point of invoice and payment; returns ready by the 15th"],
        ["Withholding tax credit notes chased manually, if at all",
         "WHT receivable register that reconciles credits and offsets them against CIT"],
        ["Project margin known only at year-end",
         "Live margin by estate and by plot, with cost-to-complete flagged"],
        ["Board pack assembled by the accountant over weeks",
         "One-page dashboard on the 5th working day"],
        ["Audit file put together for the auditor",
         "Audit file exported from the system, with a complete transaction trail"],
        ["Knowledge held by one or two people",
         "Documented SOPs, role-based access, and a system the team is trained on"],
    ],
    widths=[8.2, 8.2], size=8.9, zebra=True)

spacer(doc, 10)

h2(doc, "What the Board gets, in plain terms", space_before=2)
bullet(doc, "a number you can trust on any morning, not just in March.", bold_lead="Certainty: ")
bullet(doc, "₦774.6 million of subscriber balances with ageing, owners and follow-up dates.",
       bold_lead="Cash discipline: ")
bullet(doc, "\u20A62.5 million, with 40% on signature, 30% on configuration sign-off and 30% on go-live. "
            "No hourly surprises.", bold_lead="A fixed price: ")
bullet(doc, "no dependency on any single individual \u2014 including us.", bold_lead="Continuity: ")

spacer(doc, 4)

pull_quote(doc, "A spreadsheet records history. A system runs a business.",
           "The difference, in one line")

callout(doc, "Why Zoho Books, and not a bigger ERP",
        "Zoho Books is right-sized for a \u20A6112 million contractor: real projects, real inventory, real tax "
        "codes, at a subscription cost that will not embarrass the finance budget. It is also the system your "
        "team can be trained on in two days rather than two quarters \u2014 which is the single biggest reason "
        "implementations succeed or quietly die.")

page_break(doc)

# ----------------------------------------------------------------------------
# 5. SERVICE 2 — TAX
# ----------------------------------------------------------------------------
section_banner(doc, "05  ·  Service two", "Tax compliance & management",
               "Every filing, every deadline, every credit \u2014 owned by us, visible to you, priced monthly.")

para(doc, "Baay\u2019s tax exposure is not unusual for a fast-growing developer; it is simply unmanaged. "
          "Land sales are VAT-exempt while construction and consultancy are taxable. Clients deduct WHT from "
          "progress payments and the credit notes do not always reach the accounts. And since 1 January 2026 "
          "the consolidated Nigeria Tax Act and Nigeria Tax Administration Act have changed the filing "
          "landscape again. None of this is fatal. All of it is avoidable.", space_after=6)

h2(doc, "The compliance calendar we will own", space_before=4)

data_table(doc,
    ["Obligation", "Due", "What we do each period"],
    [
        ["VAT return & payment", "21st monthly", "Output VAT on taxable construction and consultancy; exempt "
         "land supplies excluded and documented; input VAT recovered; FIRS filing and remittance evidence"],
        ["WHT on vendors", "21st monthly", "Correct rate applied (2% building, 5% other services under the 2024 "
         "regulations), credit notes issued, schedule filed"],
        ["PAYE & pension", "10th monthly", "Payroll review, consolidated relief allowance, LIRS filing, "
         "remittance, employee schedules"],
        ["CIT self-assessment", "30 June annually", "Assessable profit, disallowables, capital allowances limited "
         "to 66.67%, minimum tax test, deferred tax under IAS 12"],
        ["TET, PTF, NASENI", "30 June annually", "3% tertiary education tax, 0.005% Police Trust Fund, 0.25% "
         "NASENI where turnover exceeds \u20A6100m"],
        ["WHT credits", "Continuous", "Reconcile client deductions to credit notes and offset against CIT; "
         "recover what has already been paid"],
        ["Related party & TP", "Annual", "Director loan disclosure (\u20A67.5m interest-free) and transfer "
         "pricing documentation"],
        ["FIRS / LIRS queries", "As they arise", "Assessments, objections, waiver requests and audit support \u2014 "
         "we sit in the meeting, not just send the letter"],
    ],
    widths=[3.6, 2.4, 10.4], first_col_bold=True, size=8.8)

spacer(doc, 4)

h2(doc, "The first 90 days: a tax health check, not a lecture", space_before=0)
para(doc, "Before we file anything new, we reconcile what has already happened \u2014 FY2023 to FY2025 \u2014 "
          "against your audited accounts and the master workbook.", space_after=5)
bullet(doc, "reconcile every tax return to the audited figures, explain every variance, and quantify exposure "
            "before FIRS does.", bold_lead="Returns vs accounts: ")
bullet(doc, "claim the \u20A61.3 million-plus already deducted by clients, and fix the control so it does not "
            "recur.", bold_lead="WHT credits: ")
bullet(doc, "bring PAYE, pension and WHT on vendors up to date, with a clean schedule for each authority.",
       bold_lead="Payroll and vendor taxes: ")
bullet(doc, "separate exempt land sales from taxable construction revenue, invoice by invoice, so you stop "
            "over-declaring VAT.", bold_lead="VAT mapping: ")
bullet(doc, "capital allowances on \u20A63.6 million of 2025 additions, and a deferred tax position that ties "
            "to the accounts.", bold_lead="Allowances and deferral: ")

spacer(doc, 2)

callout(doc, "Our promise on deadlines",
        "If a filing is late because of us, that month\u2019s fee is waived and any penalty or interest arising "
        "from our delay is borne by us. We would rather put that in writing than in a brochure.", size=9.7)

page_break(doc)

# ----------------------------------------------------------------------------
# 6. FEES
# ----------------------------------------------------------------------------
h1(doc, "06", "Fees, and what they buy")

para(doc, "Two services, priced separately so the Board can approve them separately. All figures are in "
          "Nigerian Naira and exclude VAT at 7.5%.", space_after=10)

# Fee hero table
t = make_table(doc, 3, 3, [6.2, 5.1, 5.1])
table_borders(t, color=LINE, sz="4")
header_row(t, ["Service", "Fee", "Structure"])
rows = [
    ("Zoho Books implementation", "\u20A62,500,000", "Fixed fee \u00B7 8 weeks"),
    ("Tax health check (FY2023\u20132025)", "\u20A61,500,000", "One-off \u00B7 waived on Option B"),
]
for i, (a, b, c) in enumerate(rows):
    row = t.rows[i + 1]
    fill = CREAM_2 if i % 2 else None
    for j, val in enumerate((a, b, c)):
        cell = row.cells[j]
        if fill:
            cell_bg(cell, fill)
        body_cell(cell, val, size=9.4, bold=(j in (0, 1)),
                  color=(FOREST if j == 0 else (GOLD if j == 1 else GREY)))

spacer(doc, 10)

h2(doc, "Tax compliance retainer \u2014 choose one", space_before=2)

data_table(doc,
    ["What you get", "Option A \u00B7 Compliance core", "Option B \u00B7 Compliance + advisory (recommended)"],
    [
        ["Monthly fee", "\u20A6250,000", "\u20A6400,000"],
        ["Annual equivalent", "\u20A63,000,000", "\u20A64,800,000"],
        ["VAT, WHT, PAYE", "Preparation, filing, remittance", "Preparation, filing, remittance"],
        ["CIT, TET, PTF, NASENI", "Annual return and schedules", "Annual return, planning and position paper"],
        ["Tax health check", "Billed separately at \u20A61,500,000", "Included \u2014 fee waived"],
        ["Quarterly tax review", "\u2014", "Board-ready review with cash tax forecast"],
        ["FIRS / LIRS audit support", "Per diem, quoted", "Included, up to 10 days a year"],
        ["Contract & JV tax advice", "\u2014", "Included for up to two mandates a quarter"],
    ],
    widths=[4.0, 5.6, 6.8], first_col_bold=True, size=8.9,
    fills=[None, None, None, None, None, None, None, None])

spacer(doc, 6)

para(doc, "Year one, Option B: \u20A62,500,000 + \u20A64,800,000 = \u20A67,300,000, all in \u2014 the health check "
          "is absorbed.", size=9.6, italic=True, color=FOREST, space_after=8)

h2(doc, "Payment terms and exclusions", space_before=2)
bullet(doc, "40% on engagement letter, 30% on configuration sign-off, 30% on go-live. Invoiced in naira, "
            "payable within 14 days.", bold_lead="Zoho Books: ")
bullet(doc, "monthly in advance, starting with the first filing cycle after signature.", bold_lead="Retainer: ")
bullet(doc, "\u20A6750,000 on signature, \u20A6750,000 on delivery of the report (waived under Option B).",
       bold_lead="Health check: ")
bullet(doc, "Zoho subscription paid directly by Baay to Zoho; statutory taxes, penalties and interest arising "
            "before this engagement; external legal counsel; travel outside Lagos.",
       bold_lead="Not included: ")
bullet(doc, "if FIRS or a state authority raises an assessment on a period we managed, we handle the response "
            "within the retainer.", bold_lead="Included by default: ")

spacer(doc, 6)

callout(doc, "The commercial case in one line",
        "The Zoho build costs less than six months of the manual effort it removes \u2014 and the tax retainer is "
        "covered by the \u20A61.3 million of WHT credits we expect to recover in the first quarter, before counting "
        "a single penalty avoided.")

page_break(doc)

# ----------------------------------------------------------------------------
# 7. WHY US / TEAM / GOVERNANCE
# ----------------------------------------------------------------------------
h1(doc, "07", "Why this team, and how we will behave")

h2(doc, "We already know the file", space_before=2)
para(doc, "Most implementation partners spend the first month discovering what Baay does. We have already "
          "mapped your nine estates, your four banks, your IFRS 15 revenue model, your tax computations and "
          "the 12 control checks your workbook passes. That familiarity is worth three to four weeks of fees "
          "and, more importantly, it removes the biggest migration risk: rebuilding a chart of accounts that "
          "does not match the audited statements.", space_after=8)

h2(doc, "The team", space_before=2)
data_table(doc,
    ["Role", "What they own", "Commitment"],
    [
        ["Engagement partner (FCA)", "Board relationship, sign-off, tax positions and FIRS escalations",
         "Fortnightly review"],
        ["Zoho implementation lead", "Configuration, migration, integration, training, hypercare",
         "Full-time, weeks 1\u20138"],
        ["Chartered tax manager", "VAT, WHT, PAYE, CIT and TET filings; health check; authority correspondence",
         "Monthly cycle"],
        ["Data & reporting analyst", "Bank data, master data cleansing, dashboard build", "Weeks 2\u20137"],
        ["Quality reviewer", "Independent check of the trial balance, tax schedules and SOPs before sign-off",
         "Two gates"],
    ],
    widths=[4.4, 9.0, 3.0], first_col_bold=True, size=8.9)

spacer(doc, 8)

h2(doc, "How we will work", space_before=2)
bullet(doc, "a named person, reachable on WhatsApp and email, who knows your file.",
       bold_lead="One point of contact: ")
bullet(doc, "a two-page note every Friday during implementation, and a one-page tax position every month "
            "after that.", bold_lead="Weekly visibility: ")
bullet(doc, "your accountant and client-relations team are trained alongside us, not after us. If we do our "
            "job, you will not need us for routine work.", bold_lead="Capability transfer: ")
bullet(doc, "we do not sign off our own work. A second reviewer clears the trial balance and every tax "
            "position before it reaches the Board.", bold_lead="Independent review: ")

spacer(doc, 8)

callout(doc, "The 90-day commitment",
        "Within 90 days of signature: Zoho Books live with 2026 data reconciled \u00B7 every statutory filing "
        "current \u00B7 the 2023\u20132025 tax position documented with exposures quantified \u00B7 and a board "
        "dashboard the directors can read in five minutes. If we miss those four, the next month is on us.")

# ----------------------------------------------------------------------------
# 8. NEXT STEPS / ACCEPTANCE
# ----------------------------------------------------------------------------
page_break(doc)
h1(doc, "08", "Next steps")

data_table(doc,
    ["#", "Step", "Owner", "When"],
    [
        ["1", "Board decision on this proposal", "Directors", "By 31 October 2026"],
        ["2", "45-minute scoping call to confirm access, users and estate list", "Both", "Within 5 working days"],
        ["3", "Engagement letter and data protection addendum signed", "Both", "Before kick-off"],
        ["4", "Kick-off and diagnostic note", "Us", "Within 5 working days of signature"],
        ["5", "Chart of accounts and implementation plan approved", "Baay finance", "Week 2"],
        ["6", "Go-live and first live month-end", "Us", "Week 8"],
    ],
    widths=[1.0, 8.6, 3.0, 3.8], size=9.0)

spacer(doc, 8)

h2(doc, "Assumptions behind the fees", space_before=0)
bullet(doc, "eight weeks assumes access to the four bank accounts and the client register within five working "
            "days of kick-off.", bold_lead="Timeline: ")
bullet(doc, "one company, one currency, nine estates. A second entity or a new estate is scoped separately.",
       bold_lead="Scope: ")
bullet(doc, "Zoho Books Nigeria edition, subscription paid by Baay directly to Zoho.",
       bold_lead="Platform: ")
bullet(doc, "all fees hold for 90 days from signature and are reviewed annually against transaction volume.",
       bold_lead="Fee validity: ")

spacer(doc, 8)

h2(doc, "What we need from you", space_before=0)
bullet(doc, "read-only access to the four bank accounts, and statements from January 2026 to date.",
       bold_lead="Bank access: ")
bullet(doc, "a named finance lead, two days a week for eight weeks, with authority to approve the chart of "
            "accounts.", bold_lead="One internal owner: ")
bullet(doc, "the client register, project contracts, prior tax returns and receipts — most of which we have "
            "already catalogued.", bold_lead="Documents: ")

spacer(doc, 10)

# signature block
t = make_table(doc, 1, 2, [8.0, 8.0])
no_borders(t)
for i, (who, lines) in enumerate([
    ("Accepted for Baay Projects Limited", ["Name:", "Designation:", "Signature & date:"]),
    ("For [Firm Name]", ["Name:", "Designation:", "Signature & date:"]),
]):
    cell = t.cell(0, i)
    cell_margins(cell, top=160, bottom=160, left=0 if i == 0 else 300, right=300)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(who.upper())
    style_run(r, size=7.8, bold=True, color=GOLD, spacing=1.2)
    for ln in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(14)
        r = p.add_run(ln)
        style_run(r, size=8.6, color=MID_GREY)

spacer(doc, 10)
para(doc, "This proposal is valid for 30 days from the date of issue. All figures are drawn from the audited "
          "financial statements for 2023\u20132025 and the Baay Projects master workbook, and are stated in "
          "Nigerian Naira. Final scope, deliverables and terms will be set out in the engagement letter.",
     size=7.8, italic=True, color=MID_GREY, space_after=2)
para(doc, "Baay Projects Limited \u00B7 RC 1526224 \u00B7 TIN 21548976-0001 \u00B7 Confidential",
     size=7.8, color=MID_GREY)


# ----------------------------------------------------------------------------
# OOXML housekeeping: enforce schema element order
# ----------------------------------------------------------------------------
_ORDER = {
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
             "textDirection", "tcFitText", "vAlign", "hideMark"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
              "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
              "tblLook", "tblCaption", "tblDescription"],
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
            "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
            "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
            "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"],
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
            "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
            "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
            "sectPr", "pPrChange"],
    "trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight",
             "tblHeader", "tblCellSpacing", "jc", "hidden"],
}
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _dedupe(el):
    """Drop duplicate singleton children (e.g. two tblLayout elements)."""
    seen = set()
    for child in list(el):
        tag = child.tag
        if tag in seen:
            el.remove(child)
        else:
            seen.add(tag)


def normalize_order(document):
    for local, order in _ORDER.items():
        rank = {n: i for i, n in enumerate(order)}
        for el in document.element.body.iter(_W + local):
            _dedupe(el)
            children = list(el)
            children.sort(key=lambda c: rank.get(c.tag.replace(_W, ""), len(order)))
            for c in children:
                el.remove(c)
            for c in children:
                el.append(c)
    # headers / footers too
    for sec in document.sections:
        for part in (sec.header, sec.footer):
            for local, order in _ORDER.items():
                rank = {n: i for i, n in enumerate(order)}
                for el in part._element.iter(_W + local):
                    _dedupe(el)
                    children = list(el)
                    children.sort(key=lambda c: rank.get(c.tag.replace(_W, ""), len(order)))
                    for c in children:
                        el.remove(c)
                    for c in children:
                        el.append(c)


# ----------------------------------------------------------------------------
settings = doc.settings.element
if settings.find(qn("w:updateFields")) is None:
    settings.append(_el("w:updateFields", **{"w:val": "true"}))

normalize_order(doc)
out = "BAAY_PROJECTS_LIMITED_Engagement_Proposal_Zoho_Books_and_Tax.docx"
doc.save(out)
print("saved", out)
