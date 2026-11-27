"""Withholding tax rates setup for Baay Projects vendors.

One document: the law in short, the rates that follow from Baay's
sector, the exact list to type into Zoho Books Premium, and the
screens for each step. Screenshots are Zoho's own help screens.
Sample names, dates and currency in those screens are Zoho's, not Baay's.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image as PILImage
from reportlab.lib.colors import Color, white
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    CondPageBreak,
    Flowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parents[1]
SCREENS = Path(__file__).resolve().parent / "guide-screens"
DOCX = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Withholding_Tax_Setup_Guide.docx"
PDF = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Withholding_Tax_Setup_Guide.pdf"

NAVY = "0B1F33"
TEAL = "0E7C66"
INK = "1C2833"
MUTED = "5E6A72"
WARN = "8C4A2F"
DATE = "8 October 2026"


def rgb(value):
    value = value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def rl(value):
    value = value.lstrip("#")
    return Color(int(value[0:2], 16) / 255, int(value[2:4], 16) / 255, int(value[4:6], 16) / 255)


def shade(paragraph, hex_colour):
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_colour)
    p_pr.append(shd)


def set_run(run, size, bold=False, colour=INK, italic=False):
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = rgb(colour)


def add_page_number(paragraph):
    run = paragraph.add_run()
    set_run(run, 8, colour=MUTED)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


class WordGuide:
    def __init__(self):
        self.doc = Document()
        self.fig = 0
        section = self.doc.sections[0]
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(1.7)
        section.right_margin = Cm(1.7)
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.7)
        self.usable = 17.6
        header = section.header.paragraphs[0]
        header.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = header.add_run("Baay Projects Limited")
        set_run(run, 9, True, NAVY)
        run = header.add_run("    ·    Withholding tax rates setup, vendors")
        set_run(run, 9, colour=MUTED)
        footer = section.footer.paragraphs[0]
        run = footer.add_run("Vendor withholding tax only. PAYE and VAT are set up elsewhere.    ")
        set_run(run, 8, colour=MUTED)
        add_page_number(footer)

    def _p(self, text, size=11, bold=False, colour=INK, before=2, after=6, italic=False):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.line_spacing = 1.08
        run = p.add_run(text)
        set_run(run, size, bold, colour, italic)
        return p

    def h1(self, text):
        self._p(text, 16, True, NAVY, 14, 4)

    def h2(self, text):
        self._p(text, 13, True, TEAL, 10, 2)

    def p(self, text):
        self._p(text)

    def note(self, text):
        p = self._p("Note.  " + text, 10.5, italic=True, colour="3D3A2E", before=4, after=6)
        shade(p, "F4F1E4")

    def stop(self, text):
        p = self._p("Do not.  " + text, 10.5, italic=True, colour=WARN, before=4, after=6)
        shade(p, "FBF4EE")

    def bullets(self, items):
        for item in items:
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.7)
            p.paragraph_format.first_line_indent = Cm(-0.35)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run("•  ")
            set_run(run, 11, colour=TEAL)
            run = p.add_run(item)
            set_run(run, 11)

    def step(self, number, title, body):
        self._p(f"Step {number}.  {title}", 11, True, NAVY, 8, 1)
        self._p(body, 11, after=3)

    def table(self, headers, rows, widths):
        t = self.doc.add_table(rows=1 + len(rows), cols=len(headers))
        t.style = "Table Grid"
        for i, h in enumerate(headers):
            cell = t.rows[0].cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(h)
            set_run(run, 8.5, True, "FFFFFF")
            shade(p, NAVY)
            cell.width = Cm(widths[i])
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                cell = t.rows[r + 1].cells[c]
                cell.text = ""
                p = cell.paragraphs[0]
                run = p.add_run(value)
                set_run(run, 8)
                if r % 2:
                    shade(p, "F4F7F6")
                cell.width = Cm(widths[c])
        self._p("", after=4)

    def figure(self, key, caption):
        self.fig += 1
        path = SCREENS / key
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        with PILImage.open(path) as im:
            w, h = im.size
        disp_w = self.usable
        disp_h = disp_w * h / w
        if disp_h > 11.0:
            disp_h = 11.0
            disp_w = 11.0 * w / h
        p.add_run().add_picture(str(path), width=Cm(disp_w))
        cap = self.doc.add_paragraph()
        cap.paragraph_format.space_before = Pt(1)
        cap.paragraph_format.space_after = Pt(8)
        run = cap.add_run(
            f"Figure {self.fig}.  {caption}  Zoho Books screen. Sample names, dates and currency are Zoho’s, not Baay’s."
        )
        set_run(run, 8, italic=True, colour=MUTED)

    def save(self):
        self.doc.save(DOCX)


pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))


class PdfImage(Flowable):
    def __init__(self, path, caption, number):
        super().__init__()
        self.path = str(path)
        self.caption = caption
        self.number = number
        with PILImage.open(path) as im:
            self.px_w, self.px_h = im.size
        self.dw = self.dh = 0
        self.lines = []

    def wrap(self, aw, ah):
        text = f"Figure {self.number}.  {self.caption}  Zoho Books screen. Sample data, not Baay."
        words = text.split()
        lines, cur = [], ""
        for word in words:
            trial = (cur + " " + word).strip()
            if pdfmetrics.stringWidth(trial, "DejaVu", 7.5) <= aw:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = word
        if cur:
            lines.append(cur)
        self.lines = lines
        cap_h = 3.3 * mm * max(1, len(lines)) + 2 * mm
        img_h = min(108 * mm, aw * self.px_h / self.px_w)
        img_w = img_h * self.px_w / self.px_h
        if img_w > aw:
            img_w = aw
            img_h = aw * self.px_h / self.px_w
        self.dw = img_w
        self.dh = img_h
        self.width = aw
        self.height = self.dh + cap_h
        return self.width, self.height

    def draw(self):
        self.canv.drawImage(self.path, 0, self.height - self.dh, self.dw, self.dh, preserveAspectRatio=True, mask="auto")
        self.canv.setFillColor(rl(MUTED))
        self.canv.setFont("DejaVu", 7.5)
        y = self.height - self.dh - 3.2 * mm
        for line in self.lines:
            self.canv.drawString(0, y, line)
            y -= 3.3 * mm


class PdfGuide:
    def __init__(self):
        self.story = []
        self.fig = 0
        self.s = {
            "h1": ParagraphStyle("h1", fontName="DejaVuBold", fontSize=13, leading=16, textColor=rl(NAVY), spaceBefore=10, spaceAfter=4),
            "h2": ParagraphStyle("h2", fontName="DejaVuBold", fontSize=11, leading=14, textColor=rl(TEAL), spaceBefore=8, spaceAfter=2),
            "body": ParagraphStyle("body", fontName="DejaVu", fontSize=9, leading=12.2, textColor=rl(INK), alignment=TA_JUSTIFY, spaceAfter=6),
            "note": ParagraphStyle("note", fontName="DejaVu", fontSize=8.5, leading=11.2, textColor=Color(0.24, 0.23, 0.18), spaceBefore=2, spaceAfter=4),
            "stop": ParagraphStyle("stop", fontName="DejaVu", fontSize=8.5, leading=11.2, textColor=rl(WARN), spaceBefore=2, spaceAfter=4),
            "bullet": ParagraphStyle("bullet", fontName="DejaVu", fontSize=9, leading=12, textColor=rl(INK), leftIndent=10, spaceAfter=1),
            "step": ParagraphStyle("step", fontName="DejaVuBold", fontSize=9.5, leading=12, textColor=rl(NAVY), spaceBefore=6, spaceAfter=1),
            "th": ParagraphStyle("th", fontName="DejaVuBold", fontSize=7.4, leading=9.2, textColor=white),
            "td": ParagraphStyle("td", fontName="DejaVu", fontSize=7.3, leading=9.2, textColor=rl(INK)),
        }

    def h1(self, text):
        self.story.append(Paragraph(text, self.s["h1"]))

    def h2(self, text):
        self.story.append(Paragraph(text, self.s["h2"]))

    def p(self, text):
        self.story.append(Paragraph(text.replace("&", "&amp;"), self.s["body"]))

    def note(self, text):
        self.story.append(Paragraph("Note.  " + text.replace("&", "&amp;"), self.s["note"]))

    def stop(self, text):
        self.story.append(Paragraph("Do not.  " + text.replace("&", "&amp;"), self.s["stop"]))

    def bullets(self, items):
        for item in items:
            self.story.append(Paragraph("•  " + item.replace("&", "&amp;"), self.s["bullet"]))
        self.story.append(Spacer(1, 1.5 * mm))

    def step(self, number, title, body):
        self.story.append(Paragraph(f"Step {number}.  {title}", self.s["step"]))
        self.story.append(Paragraph(body.replace("&", "&amp;"), self.s["body"]))

    def table(self, headers, rows, widths):
        data = [[Paragraph(h, self.s["th"]) for h in headers]]
        for row in rows:
            data.append([Paragraph(c.replace("&", "&amp;"), self.s["td"]) for c in row])
        total = sum(widths)
        col = [178 * mm * w / total for w in widths]
        t = Table(data, colWidths=col, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), rl(NAVY)),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, rl("F4F7F6")]),
            ("GRID", (0, 0), (-1, -1), 0.3, rl("D5DEDB")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        self.story.append(t)
        self.story.append(Spacer(1, 2 * mm))

    def figure(self, key, caption):
        self.fig += 1
        self.story.append(CondPageBreak(120 * mm))
        self.story.append(PdfImage(SCREENS / key, caption, self.fig))
        self.story.append(Spacer(1, 2 * mm))

    def save(self):
        def on_page(canvas, doc):
            canvas.saveState()
            canvas.setFillColor(rl(NAVY))
            canvas.rect(0, A4[1] - 11 * mm, A4[0], 11 * mm, fill=1, stroke=0)
            canvas.setFillColor(white)
            canvas.setFont("DejaVuBold", 8)
            canvas.drawString(16 * mm, A4[1] - 7 * mm, "Baay Projects Limited")
            canvas.setFont("DejaVu", 8)
            canvas.drawRightString(A4[0] - 16 * mm, A4[1] - 7 * mm, "Withholding tax setup")
            canvas.setFillColor(rl(TEAL))
            canvas.rect(0, 0, A4[0], 8 * mm, fill=1, stroke=0)
            canvas.setFillColor(white)
            canvas.setFont("DejaVu", 7.5)
            canvas.drawString(16 * mm, 3 * mm, "Vendor withholding tax only. PAYE and VAT are set up elsewhere.")
            canvas.drawRightString(A4[0] - 16 * mm, 3 * mm, str(doc.page))
            canvas.restoreState()

        doc = SimpleDocTemplate(
            str(PDF), pagesize=A4,
            leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=14 * mm,
            title="Baay Projects Limited — Zoho Books withholding tax setup",
            author="Baay Projects Limited",
        )
        doc.build(self.story, onFirstPage=on_page, onLaterPages=on_page)


def write(doc):
    doc.h1("What this guide is for")
    doc.p("This is the one-time configuration of withholding tax on Baay Projects vendors in Zoho Books Premium. It is not the setup guide, the posting guide or the training guide. The setup guide is the chart and the organisation settings. The posting guide is the day-to-day desk note. The training guide is what the consultant says in the room. This note is the rate list, the screens where each rate is typed, and the law those rates come from.")
    doc.p("Finance configures. The Head of Accounts approves the rate list before the first real payment, approves every override, and signs the test run. Sales and Client Relations do not open Settings > Taxes. They never need to.")
    doc.table(
        ["Person", "What they do with this guide", "What they do not do"],
        [
            ["Finance and Accounts", "Enable the feature, create the rates, tag the vendor cards, apply rates on bills, record the deduction at payment", "Invent a rate that is not in the table, or change an approved rate"],
            ["Head of Accounts", "Approve the rate list, each override, and the monthly remittance", "Delegate the approval to the person who typed the rate"],
            ["Sales, Client Relations", "Read only. They will see a shorter bill screen once vendors are tagged", "Open Settings > Taxes, or edit a vendor's tax box"],
        ],
        [3.6, 7.6, 6.4],
    )
    doc.p("Premium is already on, from the setup guide. Without Premium the withholding tax settings page does not appear. Currency is NGN. The chart is the 136-account chart of the same date as this note.")

    doc.h1("1.  The law this list follows")
    doc.p("Two instruments set the rates. The Nigeria Tax Act 2025, in force from 1 January 2026, is the income tax act under which withholding works. The Deduction of Tax at Source (Withholding) Regulations 2024, in force since 1 January 2025, carry the rate schedule used below. Where a tax adviser or a later amendment disagrees with this table, the adviser wins and this table is reissued.")
    doc.p("Baay is the withholding agent on vendor payments. Three duties follow. Deduct when the money leaves, not when the bill is dated. Remit by the 21st of the month after the deduction, to the authority the schedule names — federal for company payees, state for individual payees such as landlords and directors. Keep the schedule so each vendor can prove what was withheld.")
    doc.p("The deduction is the vendor's advance income tax, not Baay's expense. In the chart it sits in 2160 Withholding tax payable until it is remitted. The expense on a bill is always the gross amount. WHT is never netted against VAT in 2150, never parked in 1158 Advance Tax, and never taken to income.")
    doc.table(
        ["Rule", "What it means at Baay"],
        [
            ["Baay is not a small company", "The Regulations exempt a small payer (turnover at or below ₦25 million) from deducting, and only where the supplier has a valid TIN and the month's invoices to that supplier are ₦2,000,000 or less. Baay is far above that scale. Deduct on every eligible payment, of any size."],
            ["No TIN on the vendor card", "For goods, services, construction and other active-business lines the rate doubles. Use the doubled rate in the list below, or the purchases override, and tell the Head of Accounts before the payment goes out."],
            ["Failure to deduct", "Penalty of 40 per cent of the amount that should have been withheld, plus interest, under the Nigeria Tax Administration Act 2025."],
            ["Late remittance", "Penalty of 10 per cent a year on the unremitted tax, plus interest at the CBN monetary policy rate."],
            ["Treaty rates", "Only where a ratified double taxation agreement reduces the rate for a non-resident payee. Do not apply a treaty rate from a vendor's email. The Head of Accounts asks for the treaty article and the resident certificate."],
        ],
        [4.4, 13.2],
    )

    doc.h1("2.  Baay's sector, and which rates that pulls in")
    doc.p("Baay develops estates: land allocated for plots, houses built and handed over, co-ownership contracts, retainers from subscribers. The vendors are therefore contractors, material suppliers, professional firms, agents, landlords and general service providers. The rate follows what is being bought and who is being paid, not a single rate for the company. One bill can carry two rates on two lines.")
    doc.table(
        ["What Baay pays for", "Who is paid", "Chart line", "Rate to use", "Rate"],
        [
            ["Houses, roads, drains and other building works on an estate", "Resident contractor with an RC number", "1220 while building; 5010 or 2010 when released", "Construction - Building and Civil Works [Resident]", "2%"],
            ["Fencing, borehole, fit-out, structural repairs", "Resident contractor or tradesman", "6025 or 1220", "Other Construction and Related [Resident]", "5%"],
            ["Cement, tiles, rods and other materials from a distributor", "Resident supplier who is not the maker", "5010 or 1220", "Goods and Materials [Resident]", "2%"],
            ["The same goods bought direct from the manufacturer", "Manufacturer", "5010", "None — exempt. Note it on the bill", "0%"],
            ["Legal, audit, tax, engineering, architecture, project management", "Resident firm, RC or BN", "6030; audit and tax advice is 6040", "Professional and Consultancy Fees [Resident]", "5%"],
            ["Security, cleaning, haulage and other general services", "Resident service provider", "6090, 6095", "General Services [Resident]", "2%"],
            ["Selling commission, including an upline's share", "Resident agent", "6035, gross", "Commission and Brokerage [Resident]", "5%"],
            ["Office rent, site rent, staff lodge, hire of equipment", "Landlord or lessor", "6020, 6065, 5030", "Rent, Hire and Lease", "10%"],
            ["Directors' fees", "A director, individual", "Remuneration lines", "Directors' Fees [Resident Individual]", "15%"],
            ["Interest on a loan", "Lender", "7010 finance cost", "Interest and Dividends [Corporate Resident]", "10%"],
            ["A consultant outside Nigeria", "Non-resident", "6030", "Professional Fees [Non-Resident]", "10%, final for the vendor"],
            ["Any eligible line, vendor has no TIN", "Anyone without a TIN", "Same as the line", "The matching No TIN rate below", "Double"],
        ],
        [4.6, 3.4, 3.4, 4.6, 1.6],
    )
    doc.p("Ordinary trading stock Baay buys to resell, payments across a counter in cash or instant transfer, goods the supplier manufactured, and the fuels named in the Regulations are outside withholding. If a line is outside the schedule, no rate goes on it, and the reason goes in the bill notes.")

    doc.h1("3.  The full rate schedule, for reference")
    doc.p("First Schedule of the Deduction of Tax at Source (Withholding) Regulations 2024, the rows Baay is likely to meet. Final means the vendor pays no further tax on that income.")
    doc.table(
        ["Payment", "Company, resident", "Company, non-resident", "Individual, resident", "Individual, non-resident"],
        [
            ["Commission, consultancy, technical, management, professional fees", "5%", "10% final", "5%", "10% final"],
            ["Supply of goods or materials, not by the manufacturer", "2%", "—", "2%", "—"],
            ["Services not specifically listed", "2%", "5%", "2%", "5%"],
            ["Construction of roads, bridges, buildings, power plants", "2%", "5%", "2%", "5%"],
            ["Any other construction and related activities", "5%", "10% final", "5%", "10% final"],
            ["Brokerage fee", "5%", "10%", "5%", "10%"],
            ["Rent, hire or lease", "10%", "10%", "10%", "10%"],
            ["Dividends and interest", "10%", "10%, treaty may reduce", "10%", "10%, treaty may reduce"],
            ["Royalties", "10%", "10%", "5%", "5%"],
            ["Directors' fees", "—", "—", "15%", "20%"],
        ],
        [5.4, 3.1, 3.4, 3.0, 3.4],
    )
    doc.note("Small-company turnover for the payer exemption is read from the Regulations as ₦25 million. The Nigeria Tax Act 2025 raises the small-company band for company income tax to ₦100 million of turnover and ₦250 million of fixed assets. That wider band does not hand Baay the withholding exemption. Baay is not small under either figure.")

    doc.h1("4.  The rate list to type into Zoho")
    doc.p("Create exactly these rates. The name is what Finance will see on every bill, so type it the same way each time. Tax Type is the statutory category in Zoho's dropdown — pick the row that matches; if your organisation's list is worded differently, choose the closest and keep the full wording in the name. On every rate, Account to Track Withholding Tax in Purchases is 2160 Withholding tax payable, and Account to Track Withholding Tax in Sales is 1150 Withholding tax credit receivable. If Zoho offers its own sample account names, open the chart first and make sure Baay's 2160 and 1150 are the accounts being offered.")
    doc.table(
        ["Withholding Tax Name (type exactly)", "Rate", "Tax Type to pick", "Purchases account", "Sales account"],
        [
            ["Construction - Building and Civil Works [Resident]", "2%", "Construction of roads, bridges, buildings [Resident]", "2160", "1150"],
            ["Goods and Materials [Resident]", "2%", "Supply of goods or materials [Resident]", "2160", "1150"],
            ["General Services [Resident]", "2%", "Supply or rendering of services [Resident]", "2160", "1150"],
            ["Professional and Consultancy Fees [Resident]", "5%", "Commission, consultancy, technical, management and professional fees [Resident]", "2160", "1150"],
            ["Other Construction and Related [Resident]", "5%", "Any other construction and related activities [Resident]", "2160", "1150"],
            ["Commission and Brokerage [Resident]", "5%", "Brokerage fee [Resident]", "2160", "1150"],
            ["Rent, Hire and Lease", "10%", "Rent, hire or lease", "2160", "1150"],
            ["Interest and Dividends [Corporate Resident]", "10%", "Dividends and Interest [Corporate Resident]", "2160", "1150"],
            ["Directors' Fees [Resident Individual]", "15%", "Directors' fees [Resident Individual]", "2160", "1150"],
            ["Professional Fees [Non-Resident]", "10%", "Commission, consultancy, technical, management and professional fees [Non-Resident]", "2160", "1150"],
            ["No TIN - Services and Goods [Resident]", "4%", "Same type as the doubled line", "2160", "1150"],
            ["No TIN - Professional [Resident]", "10%", "Same type as the doubled line", "2160", "1150"],
            ["No TIN - Construction [Resident]", "4%", "Same type as the doubled line", "2160", "1150"],
        ],
        [5.8, 1.4, 5.6, 2.4, 2.4],
    )
    doc.p("An optional fourteenth rate, Exempt [Manufacturer or Export] at 0 per cent, exists only if the Head of Accounts wants a manufacturer-direct purchase to show on the bill as a decision that was considered and not withheld. It is not required. A note on the bill does the same job.")

    doc.h1("5.  Configure Zoho Books Premium")
    doc.step("1", "Turn the feature on", "Settings > Taxes, under Taxes and Compliance, then Withholding Tax Settings. Slide Enable Withholding Tax on. Enable Withholding Tax For: Vendors. This guide is the vendor side only; the customer side, which feeds 1150 when a client withholds from Baay, is a later decision of the Head of Accounts. Tick Enable Withholding Tax Override for purchases transactions, so a no-TIN doubling or a correction can be typed on the bill and approved. Leave the sales override unticked. Round Off Withholding Tax: at transaction level. Save.")
    doc.figure(
        "whs-settings.png",
        "Withholding Tax settings. Choose Vendors, not the Customers and Vendors shown in the sample. Tick the purchases override, leave the sales override alone, round off at transaction level. The sample wording is Zoho's.",
    )
    doc.note("If the settings page is missing, the organisation is not on Premium. Do not fake the feature with manual journals every month. Upgrade first, then continue.")
    doc.step("2", "Create the thirteen rates", "Still under Withholding Tax Settings, click + New Rate for each row of the table in section 4. Name, rate, Tax Type, then the two accounts: purchases to 2160, sales to 1150. Save. Work down the list in order. When the list is finished, count it: thirteen rates, or fourteen with the optional exempt rate. A missing rate is found later as a bill with no deduction.")
    doc.figure(
        "whs-new-rate.png",
        "New Withholding Tax. Type the name from the table, the rate, pick the matching Tax Type, and point Purchases at 2160 and Sales at 1150. The sample rate and type on this screen are Zoho's, not Baay's.",
    )
    doc.step("3", "Tag each vendor card", "Purchases > Vendors. Open each active vendor. Under Other Details, tick Enable Withholding Tax for this Vendor, then choose the default rate that matches what Baay usually buys from them: a building contractor gets Construction 2 per cent; a law firm gets Professional 5 per cent; a landlord gets Rent 10 per cent; a land agent gets Commission 5 per cent; a distributor gets Goods 2 per cent. Write the vendor's TIN in the tax identification field or the notes. If there is no TIN, write NO TIN in capitals in the notes — that vendor's payments need the doubled rate or an approved override.")
    doc.figure(
        "whs-vendor.png",
        "Edit Vendor, Other Details. Tick Enable Withholding Tax for this Vendor and set the default rate for what this vendor normally sells Baay. Ignore GST in the sample: Baay does not use it. The sample rate is Zoho's.",
    )
    doc.p("The default rate only pre-fills. The rate on each bill line is still checked at the bill, because one vendor can bill both a professional fee and a delivery on the same document.")
    doc.step("4", "Apply the rate on the bill, line by line", "Purchases > Bills > + New, the same bill as always, with the same Account column: 5010, 6030, 6020, 1220 and so on. On each line that suffers withholding, tick Apply Withholding Tax and pick the rate for that line. Leave the tick off lines that are exempt — manufacturer-direct goods, or a line outside the schedule. A mixed bill shows one rate on the materials line and another on the professional line, which is correct.")
    doc.figure(
        "whs-line.png",
        "The item table with Apply Withholding Tax on the line. Tick only the lines that suffer the tax, and pick the rate for that line. Sample item, currency and rates are Zoho's.",
    )
    doc.p("Below the totals the Withholding Tax section lists each rate used, the amount, and the Balance Due — what the vendor will actually receive. The bill itself does not post the deduction. Zoho holds it until the payment is recorded. Save as Open.")
    doc.figure(
        "whs-bill.png",
        "The Withholding Tax section under the total. The deduction reduces Balance Due, not the bill total. Ignore the VAT line and the sample currency in this picture: Baay's bills are in naira with no VAT on these lines.",
    )
    doc.step("5", "Record the payment, and only then the tax posts", "Open the bill. Record Payment. Paid Through is the bank that paid: 1010 Providus, 1020 First Bank, 1030 Sterling or 1040 GTBank. Set Withholding Tax Deducted to Yes. The Withholding Tax Amount comes from the bill — check it against your own sum before you save. Paid amount is the net. Save as Paid. The journal behind that button debits the bill's expense or cost account gross, credits 2160 for the deduction, and credits the bank for the net. Then match the bank withdrawal to the payment, as in the posting guide.")
    doc.figure(
        "whs-payment.png",
        "Record Payment on a bill, Withholding Tax Deducted set to Yes. The amount shown is the deduction that will credit 2160. Sample amount, date and Paid Through are Zoho's.",
    )
    doc.step("6", "Remit, and keep the proof", "At month-end, run the report on 2160. The balance is what has been deducted and not yet paid over. Remit by the 21st of the following month: company payees to the federal revenue service, individual payees — rent to a person, directors' fees — to the state revenue service. Attach the remittance receipt to the month, the same folder as the bank statements. The vendor keeps their own copy of what was withheld; that is their credit, on their side of the transaction.")
    doc.p("The customer side, where a Baay client withholds from Baay, credits 1150 when it is set up. It is not part of this configuration and is not a reason to delay the vendor list.")

    doc.h1("6.  Worked examples in naira")
    doc.p("Documentation only. These figures are not postings. Do not enter them into Zoho. They are here so Finance can check a rate by eye before the first live payment.")
    doc.table(
        ["What is being paid", "Gross", "Rate", "Withheld", "Vendor receives", "Credit"],
        [
            ["Contractor building units on Phase 1", "₦10,000,000", "2%", "₦200,000", "₦9,800,000", "2160"],
            ["Architect's fee for a block of flats", "₦2,500,000", "5%", "₦125,000", "₦2,375,000", "2160"],
            ["Security for the site office, a quarter", "₦900,000", "2%", "₦18,000", "₦882,000", "2160"],
            ["Office rent for the year", "₦1,200,000", "10%", "₦120,000", "₦1,080,000", "2160"],
            ["Land agent's commission on a sale", "₦1,500,000", "5%", "₦75,000", "₦1,425,000", "2160"],
            ["Tiles from a distributor, not the factory", "₦3,000,000", "2%", "₦60,000", "₦2,940,000", "2160"],
            ["Doors bought direct from the manufacturer", "₦5,000,000", "Exempt", "₦0", "₦5,000,000", "None"],
            ["Plumbing repair, vendor has no TIN", "₦400,000", "4% (double 2%)", "₦16,000", "₦384,000", "2160"],
            ["Consultant outside Nigeria", "₦2,000,000", "10% final", "₦200,000", "₦1,800,000", "2160"],
        ],
        [5.6, 2.6, 2.4, 2.4, 2.8, 1.8],
    )
    doc.p("If any row disagrees with what the bill screen calculates, stop. The rate on the card or the line is wrong, not the law table. Fix the rate, do not override the amount to force agreement, unless the Head of Accounts has approved that specific override in writing on the bill.")

    doc.h1("7.  Test run before the first real payment")
    doc.p("Finance does this. The Head of Accounts watches and signs. Twenty minutes. Nothing from the examples above is entered.")
    doc.bullets([
        "Create a vendor card named Test Vendor. Enable withholding tax. Default rate: General Services [Resident] 2 per cent. Notes: TIN 00000000-0000.",
        "Raise a bill on Test Vendor for ₦100,000, Account 6090, apply General Services 2 per cent on the line. Save as Open. Confirm Balance Due reads ₦98,000.",
        "Record Payment from Providus, Withholding Tax Deducted: Yes. Amount should be ₦2,000. Save as Paid.",
        "Open the journal behind that payment. It must read: debit 6090 ₦100,000, credit 2160 ₦2,000, credit 1010 ₦98,000. If 2160 is missing, the rate's Purchases account is wrong. Fix the rate, delete the payment, and run the test again.",
        "Delete the payment and the bill. Set Test Vendor inactive. Confirm 2160 is nil.",
    ])
    doc.stop("Leave the test bill, the test payment or Test Vendor active in any form. The practice is finished only when 2160 is nil and the vendor list contains no test card.")

    doc.h1("8.  What not to do")
    doc.table(
        ["Wrong", "Right"],
        [
            ["Expense the deduction, or reduce the bill", "The bill stays gross. Only the payment is net. The credit sits in 2160 until remitted"],
            ["Park a vendor deduction in 1158 Advance Tax, or net it in 2150 with VAT", "Vendor deductions are 2160. Customer-side credits are 1150. VAT is untouched"],
            ["Wait for the invoice date, or for month-end, to deduct", "Deduct when the payment leaves the bank"],
            ["Take a vendor's word for a treaty rate or an exemption", "The Head of Accounts asks for the document before the rate changes"],
            ["Leave a no-TIN vendor at the standard rate because the payment is small", "Double the rate, or use an approved override. Size does not remove the rule"],
            ["Let Sales or Client Relations edit the rate list", "Settings > Taxes is Finance, with the Head of Accounts approving"],
            ["Enter the worked naira examples as real bills", "They are documentation. Test Vendor is the only practice card"],
            ["Deduct on a manufacturer-direct purchase or an over-the-counter payment", "No rate. Write the reason in the bill notes"],
        ],
        [8.6, 8.6],
    )

    doc.h1("9.  Approval")
    doc.p("The rate list in section 4, the settings in section 5 and the test in section 7 are approved before the first live vendor payment. Any later change to a rate is a new approval, not an edit made quietly on a Friday.")
    doc.table(
        ["What", "Who", "Name", "Signature", "Date"],
        [
            ["Rate list and settings configured", "Finance", "", "", ""],
            ["Rate list approved", "Head of Accounts", "", "", ""],
            ["Test run passed, 2160 nil after delete", "Finance", "", "", ""],
            ["Test run witnessed", "Head of Accounts", "", "", ""],
        ],
            [6.0, 3.4, 3.4, 3.0, 1.8],
    )
    doc.p("Prepared for Baay Projects Limited. " + DATE + ". Rates from the Nigeria Tax Act 2025 and the Deduction of Tax at Source (Withholding) Regulations 2024. Account codes from the chart of the same date. This note is configuration guidance, not a tax ruling; a treaty case or a novel line goes to Baay's tax adviser before the payment, not after.")


def main():
    needed = [
        "whs-settings.png", "whs-new-rate.png", "whs-vendor.png",
        "whs-line.png", "whs-bill.png", "whs-payment.png",
    ]
    missing = [name for name in needed if not (SCREENS / name).exists()]
    if missing:
        raise SystemExit(f"Missing screenshots: {missing}")
    word = WordGuide()
    write(word)
    word.save()
    pdf = PdfGuide()
    write(pdf)
    pdf.save()
    print(DOCX, "figures", word.fig)
    print(PDF, "figures", pdf.fig)


if __name__ == "__main__":
    main()
