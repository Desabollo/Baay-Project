"""Training guide the consultant uses with Baay Projects Limited.

One document, four rooms: Sales, Client Relations, Finance and Accounts,
and the Head of Accounts. Screenshots are Zoho's own help screens.
Sample names in those screens are Zoho's, not Baay's.
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
DOCX = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Training_Guide.docx"
PDF = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Training_Guide.pdf"

NAVY = "0B1F33"
TEAL = "0E7C66"
INK = "1C2833"
MUTED = "5E6A72"
WARN = "8C4A2F"
DATE = "25 November 2026"


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
        run = header.add_run("    ·    Training guide for the consultant")
        set_run(run, 9, colour=MUTED)
        footer = section.footer.paragraphs[0]
        run = footer.add_run("Train on the test subscriber only.    ")
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

    def say(self, text):
        p = self._p("Say this.  " + text, 10.5, italic=True, colour="5A3E1B", before=4, after=6)
        shade(p, "F4E8D4")

    def stop(self, text):
        p = self._p("Stop the room.  " + text, 10.5, italic=True, colour=WARN, before=4, after=6)
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
            f"Figure {self.fig}.  {caption}  Zoho Books screen. Names, dates and currency in the picture are Zoho’s sample, not Baay’s."
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
            "say": ParagraphStyle("say", fontName="DejaVu", fontSize=8.5, leading=11.2, textColor=Color(0.35, 0.24, 0.11), spaceBefore=2, spaceAfter=4),
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

    def say(self, text):
        self.story.append(Paragraph("Say this.  " + text.replace("&", "&amp;"), self.s["say"]))

    def stop(self, text):
        self.story.append(Paragraph("Stop the room.  " + text.replace("&", "&amp;"), self.s["stop"]))

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
            canvas.drawRightString(A4[0] - 16 * mm, A4[1] - 7 * mm, "Training guide")
            canvas.setFillColor(rl(TEAL))
            canvas.rect(0, 0, A4[0], 8 * mm, fill=1, stroke=0)
            canvas.setFillColor(white)
            canvas.setFont("DejaVu", 7.5)
            canvas.drawString(16 * mm, 3 * mm, "Train on the test subscriber only")
            canvas.drawRightString(A4[0] - 16 * mm, 3 * mm, str(doc.page))
            canvas.restoreState()

        doc = SimpleDocTemplate(
            str(PDF), pagesize=A4,
            leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=14 * mm,
            title="Baay Projects Limited — Zoho Books training guide",
            author="Baay Projects Limited",
        )
        doc.build(self.story, onFirstPage=on_page, onLaterPages=on_page)


def write(doc):
    doc.h1("How to use this guide")
    doc.p("This is the script for the training day. It is not the setup guide, and it is not the posting guide. The setup guide is what you configure before anyone sits down. The posting guide is what Finance keeps on the desk afterwards. This note is what you say, what you click, and what you do not let the room click.")
    doc.p("Train in three sittings. Do not put Sales in the room while Finance is publishing a journal. People copy the last button they saw.")
    doc.table(
        ["Sitting", "Who is in the room", "How long", "What they leave able to do"],
        [
            ["1", "Everyone, for the opening only", "20 minutes", "Sign in. Say the one rule. Know which menu is theirs."],
            ["2", "Sales", "40 minutes", "Find a subscriber on their estate and read the balance. Nothing else."],
            ["3", "Client Relations", "90 minutes", "Open a subscriber, raise a retainer, record a payment, set the status."],
            ["4", "Finance and Accounts", "Two and a half hours", "Match the bank, post a bill, recognise a sale, and lock a month."],
            ["5", "Head of Accounts", "30 minutes, after Finance", "Read the five screens, and approve or refuse a lock."],
        ],
        [1.6, 4.6, 3.6, 7.4],
    )
    doc.p("Before the day, the setup sign-off is ticked. Premium is on. The chart is imported. The retainer liability is Unearned Revenue, 2020. The test customer named Test Subscriber exists on project Head Office, or you create it yourself before the room arrives. You do not create it in front of Sales.")
    doc.stop("If the retainer liability account is Sales, do not start. Every practice receipt will become income. Fix it from the setup guide, void the test, and only then bring people in.")

    doc.h1("1.  Opening, with everyone")
    doc.say("Zoho Books is where the receipt, the estate and the account have to agree. A receipt is not a sale. Land becomes a sale when the plot is allocated. A house becomes a sale when it is handed over. Co-ownership is not a sale at all. If you remember one sentence, remember that a receipt is not a sale.")
    doc.p("Have each person sign in with their own login. Two-factor authentication is already on. Do not let two people share a screen login. A shared login makes the lock useless, because you cannot see who posted.")
    doc.step("1", "Show the left-hand menu, then hide what is not theirs", "Stand at the projector on an admin login only long enough to name the menus. Then switch to a Sales login, and let them see how short their menu is. That is the lesson. The software, not a speech, is what stops a salesperson raising an invoice.")
    doc.table(
        ["Person", "Menus they use", "Menus they will not see, or must not open"],
        [
            ["Sales", "Customers, their project, the collections report", "Invoices, Banking, Bills, Manual Journals, the chart"],
            ["Client Relations", "Customers, Projects, Retainer Invoices, Payments Received", "Banking, Bills, Journals, Credit Notes, the chart"],
            ["Accountant", "Banking, Bills, Expenses, Invoices, Journals, Reports", "The chart. Unlocking a month."],
            ["Finance manager", "The accountant’s menus, plus the lock", "Changing the chart. Unlocking without the Head of Accounts."],
            ["Head of Accounts", "All of the above, and the chart", "Nothing is held back. They still do not re-key a receipt."],
        ],
        [3.4, 6.6, 7.2],
    )
    doc.p("Projects are the estates. A subscriber is a customer. A product is an item. Do not let anyone add an account because a new estate has opened. A new phase is a new project, and only Finance or the Head of Accounts adds it.")
    doc.p("Ask the room, before you split them: if ₦1,000,000 arrives today for a plot that has not been allocated, which account receives it? The answer is Unearned Revenue, 2020. If anyone says Sales, stay on this point. Do not move on to keep the timetable.")

    doc.h1("2.  Sales")
    doc.p("Sales can see the subscribers on their own estate, and what is still to be collected. They cannot take a payment, raise an invoice, or see a commission. If a salesperson asks for a wider login, the answer is no. The balance they need is on the retainer, not in the commission file.")
    doc.say("You are here to answer a subscriber who asks what they have paid and what is left. You are not here to receipt the money. Client Relations receipts it, after the bank has it. If you receipt it yourself, the books and the bank will disagree, and the subscriber will have two stories.")
    doc.h2("Find the subscriber")
    doc.step("1", "Open Customers, and search", "Sales > Customers. Type the subscriber’s name in the search box. Open the card only if the name is already there. Do not click + New. A second card for the same person splits the deposit from the later invoice, and nobody can see the balance.")
    doc.figure(
        "customers-list.png",
        "The customer list. Search before anything else. The Receivables column is not the deposit. Unused credits is not a retainer. Ignore the sample currencies, and ignore Place of Supply. Baay’s list is in naira.",
    )
    doc.step("2", "Read the card, and do not use New Transaction", "Open the subscriber. Payment terms should say Due on Receipt. The plot or unit is in the notes or the custom field, not in the customer name. The balance still to collect is the contract amount on the retainer, less what has been received. It is not the Receivables figure, and it is not Unused credits.")
    doc.figure(
        "customer-menu.png",
        "New Transaction is the trap on this screen. It offers Invoice and Customer Payment. Sales does not use either. Client Relations uses Retainer Invoice, and they start from the project, not from this button. The sample currency is dollars. Baay is naira. Do not send a portal invitation.",
    )
    doc.p("Have each salesperson find one subscriber on the estate they actually sell. Ask them to say, out loud, the estate, the plot or unit, what has been received, and what is left. If they cannot see the estate, their login is pointed at the wrong project. Fix the login later. Do not give them a second login in the room.")
    doc.h2("See what was collected this month")
    doc.step("3", "Open the collections report, already saved", "Reports, or the custom report named Subscriber funds. If they are looking at a raw Payments Received report, filter it to Retainer Payment, to their project, and to this month. An invoice payment is not a collection. A collection is a retainer payment.")
    doc.figure(
        "payments-report.png",
        "Payments Received. Two lines are marked Retainer Payment. That is the collection. The other lines are invoice payments, and Sales must not treat them as new money in. Filter to Retainer Payment. The amounts are Zoho’s sample, in dollars.",
    )
    doc.p("Ask one person to point at a line and say whether it is a collection or a sale. If they point at an invoice payment and call it a collection, go back to the filter. Do not leave the report on All.")
    doc.h2("What Sales practises, and then stops")
    doc.bullets([
        "Search for Test Subscriber. Read the balance. Close the card.",
        "Open the collections report. Say which lines are retainer payments.",
        "Do not click + New, Invoice, Customer Payment, or Record Payment.",
    ])
    doc.stop("If a salesperson has created a customer, an invoice, or a payment, void it before they leave the chair. Then take the extra permission off the login. The practice is not finished until the test organisation is back where it started.")
    doc.p("Hand them this card. It is the whole module.")
    doc.table(
        ["Sales may", "Sales may not"],
        [
            ["Search a subscriber on their own estate", "Create a customer"],
            ["Read the plot, the plan, and the balance still to collect", "Raise a retainer, an invoice, or a payment"],
            ["Open the retainer-payment report for their estate", "Open another estate’s list, or the commission file"],
            ["Tell Client Relations that money has been promised", "Tell Finance that a sale has happened, unless the plot is allocated or the house is handed over"],
        ],
        [8.6, 8.6],
    )

    doc.h1("3.  Client Relations")
    doc.p("Client Relations opens the subscriber, raises the retainer, records the payment once the money is in the bank, and sets the status. They do not categorise a bank line, publish a journal, or decide that a sale has happened. Completed on the register is not a sale. Allocation is the land sale. Handover is the housing sale. Finance posts both.")
    doc.say("You receipt what the bank has already received. You do not receipt a promise, a transfer that has not arrived, or a cheque that has not cleared. The status you set tells Finance what to look at. It does not post the sale.")
    doc.h2("Open the subscriber, or create one")
    doc.step("1", "Search first", "Sales > Customers. Type the full name. Open the card if it exists. A second card is the usual reason a balance looks unpaid when the money is already in. If the spelling on the bank narration is different, do not make a new card. Note the spelling and tell Finance. Finance parks the receipt in 2045 until the name is agreed.")
    doc.step("2", "Create the card only when the search is empty", "Click + New. A person is Individual, not Business. Display name is the subscriber’s name, as the register spells it. Currency is NGN. Payment terms are Due on Receipt. Leave the phone and the email blank unless you will actually write to that person from Zoho. Do not type a date of birth, and do not type a bank account number.")
    doc.figure(
        "new-customer.png",
        "New Customer. For a person, choose Individual, not Business. Currency is naira, not the dollars in this sample. Untick VAT MOSS and TDS. Payment terms stay Due on Receipt. Leave the phone blank unless you need it. Do not copy the sample company, Zylker.",
    )
    doc.p("On Custom Fields, set Register status to Prospect, Paying, Completed, Refund, Void or Incentive. Put the block and plot in Plot or unit. Set Payment plan to Outright, 3 Months, 6 Months, 12 Months, or Other. These are the words on the client register. Do not invent Paying-Bought-Handed over. Those words are not on the register.")
    doc.h2("Raise the retainer from the project")
    doc.step("3", "Start from the estate, not from the customer button", "Time Tracking > Projects. Open the estate, for example Green City Phase 1 (Ketu-Epe), or Baay Foreshore, or Eden Residence. New Transaction > Create Retainer Invoice. Starting here puts the estate on the document. If you start from the customer’s New Transaction menu, you can save a retainer with no project. Do not.")
    doc.figure(
        "project-menu.png",
        "New Transaction, on the project. Zoho’s article draws a box around Create Manual Journal. Ignore that box. Client Relations chooses Create Retainer Invoice. They do not choose Invoice, and they do not choose a journal.",
    )
    doc.step("4", "Fill the retainer, then stop and read it back", "Customer: the subscriber you just found, not a new name typed into the box. Project: the estate. Item: Land subscription for a plot, Housing subscription for a house or unit. The account on the line is 2020 Unearned Revenue. If it says Sales, stop. Tax: none. Rate: the amount the bank has received, not the contract price. Put the contract price in the Contract amount field. Reference: the plot or unit. Save. A draft does not receipt the money.")
    doc.figure(
        "retainer-new.png",
        "The retainer invoice. In the room, the customer is the subscriber, the project is the estate, the item is Land subscription or Housing subscription, and the account is 2020. Tax is none. The sample on this screen is Zoho’s. Do not copy its names or its currency.",
    )
    doc.p("Read the retainer back before you record the payment. Estate. Plot. Amount received. Contract amount. Balance. That is what the company’s own receipt shows, and the subscriber already expects it. A corner piece, a development levy and a documentation fee are not typed onto the retainer. They are invoice lines, later, and Finance raises them on allocation.")
    doc.h2("Record the payment, and only then")
    doc.step("5", "Record Payment from the retainer", "Open the retainer. Record Payment. Amount: what arrived, not the contract. Deposit to: the bank that received it, 1010 Providus, 1020 First Bank, 1030 Sterling or 1040 GTBank. Date: the bank date, not today if today is later. Reference: the narration, shortened, with no account number. Save.")
    doc.figure(
        "retainer-payment.png",
        "Record Payment sits on the retainer. Deposit to is the Baay bank that received the money, in naira. In this sample the document may still show a balance and a foreign currency. Yours is the amount on the statement, and the currency is naira.",
    )
    doc.p("Do not record a Customer Payment from the customer menu. That creates unused credits in receivables. A subscriber who is still paying is not a debtor. Do not categorise the bank line yourself. You have recorded the payment. Finance matches the bank line to it.")
    doc.step("6", "Set the status, and attach the receipt", "If money has been received and the plot is not yet allocated, status is Paying. Completed means the register says the price has been paid. It does not mean Finance has recognised a sale. Refund, Void and Incentive are set only when the register says so. Attach the receipt. Do not email it from Zoho unless the Head of Accounts has said that this subscriber is written to from Zoho.")
    doc.h2("The practice")
    doc.p("Each person, on the projector or on their own login, does this and then stops.")
    doc.bullets([
        "Search for Test Subscriber. Do not create a second one.",
        "From project Head Office, raise a retainer for ₦1,000. Item: Land subscription. Tax: none. Read the account. It must be 2020.",
        "Record the payment to Providus. Do not match the bank. That is Finance’s step.",
        "Set the status to Paying. Attach nothing that names a real subscriber.",
    ])
    doc.stop("Leave the ₦1,000 in place until Finance voids it in their sitting. Do not let Client Relations void it, and do not let them raise a second one to “see the button again”. One test is enough. If the credit was Sales and not 2020, stop the day.")
    doc.table(
        ["Client Relations may", "Client Relations may not"],
        [
            ["Search, and create a subscriber only if the search is empty", "Create a customer from the first word of a bank narration"],
            ["Raise a retainer from the project, for money already in the bank", "Raise an invoice, a credit note, or a default charge"],
            ["Record the payment on that retainer, to the bank that received it", "Categorise a bank line, or accept Other Income"],
            ["Set Prospect, Paying, Completed, Refund, Void or Incentive", "Treat Completed as a sale, or publish a journal"],
        ],
        [8.6, 8.6],
    )

    doc.h1("4.  Finance and Accounts")
    doc.p("This sitting is the accountant and the finance manager together. The accountant posts. The finance manager posts and locks. Neither of them changes the chart. The Head of Accounts is not in this sitting, except to watch the first lock.")
    doc.say("Client Relations has recorded what the subscriber paid. You make the bank agree with that record, and you refuse every suggestion that turns the receipt into income. You recognise a sale only on allocation or handover, and you release the cost the same day. A draft is not a posting. Publish is.")
    doc.h2("The bank, the same day")
    doc.step("1", "Match what already exists", "Banking, then the bank that received the money. Find the line. If a retainer payment already exists for that amount and that subscriber, click Match. Do not click Categorize. Match links the statement to the document. Categorize creates a new posting, and that is how a receipt becomes income.")
    doc.figure(
        "bank-match.png",
        "Match is the correct click when Client Relations has already recorded the retainer payment. Import Statement is the other path, and only when the feed has not brought the line. Do not categorise a matched line as well.",
    )
    doc.step("2", "Show the trap, and do not take it", "On a deposit Zoho will offer Other Income, and it may mark the line Recognized. That word is on the button. It is the wrong word for a Baay receipt. Have everyone in the room point at the button and say they will not click it. Then match the line, or categorise it to 2045 if the name is not known.")
    doc.figure(
        "bank-categorize.png",
        "The trap. The badge says Other Income, Recognized, and the button offered is Categorize. For a subscriber deposit, do not accept it. Match the retainer payment. If the name is unknown, categorise to 2045, not to income.",
    )
    doc.p("A transfer between Providus, First Bank, Sterling and GTBank is a transfer. Add Transaction, the transfer option, the other bank, the amount, the date it left. Match both statement lines to that one transfer. Two incomes, or an income and an expense, will make profit wrong in both directions and the cash will still look right.")
    doc.p("A receipt you cannot name is 2045, the same day. A payment you cannot name is 1180. Naming it is a weekly job, before the lock. Parking it in income so the month looks tidy is a misstatement. Have the accountant say which account they would use for a narration that only says TRF. If they hesitate, the answer is 2045 until someone has read it.")
    doc.h2("Bills, payments, and a small expense")
    doc.step("3", "Raise the bill before you pay it", "Purchases > Bills > + New. Search the vendor. Do not make a new card for a new spelling of the same supplier. The Account on the line is what you bought. A selling commission, including an upline’s share, is 6035, gross. It is not deducted from the receipt and not deducted from revenue. A staff lodge is 6065, not the cost of the house. Land still held is 1210. A house still being built is 1220.")
    doc.figure(
        "bill-account.png",
        "The column that matters is Account. Ignore the laptop, the tax and the discount in the sample. On a Baay bill that column is 6035, 6065, 6020, 1210, 1220 or 1230. Never Sales, and never an income account.",
    )
    doc.step("4", "Pay the bill from the bank that paid it", "Open the bill. Record Payment. Paid Through is 1010, 1020, 1030 or 1040. Then match the withdrawal to that payment. A bank charge taken off the payment is 6050, not a reduction of the bill. Withholding tax you deduct is not a discount. The bill is gross. The net goes to the vendor. The balance you still owe the tax authority is 2160.")
    doc.figure(
        "vendor-payment.png",
        "Record Payment on a bill. Vendor, amount, date, and Paid Through. This sample is in dollars, paid from Petty Cash. Yours is naira, and Paid Through is the bank that paid. Do not tick a short payment unless the bank really paid short.",
    )
    doc.step("5", "An expense with no bill", "Purchases > Expenses > + New, or categorise the bank line. The expense account is the thing bought. Paid Through is the bank, or 1055 if the site float paid it. Tax is none unless the Head of Accounts has said otherwise. Attach the receipt. Site imprest is not an expense when you hand the float out. Handing the float out debits 1055. Spending it credits 1055 and debits the cost.")
    doc.figure(
        "record-expense.png",
        "The expense screen. Ignore the sample: dollars, Cost of Goods Sold, and a tax rate. For Baay the currency is naira, tax is none unless you have been told otherwise, and the account is the line you named. Paid Through is the bank or 1055, not Advance Tax.",
    )
    doc.h2("When land is allocated")
    doc.p("Talk through the example. Do not post it live. A Phase 1 plot, contract ₦3,000,000, retainer already held ₦1,500,000, cost in 1210 of ₦900,000. These figures are an example. The register is the source.")
    doc.step("6", "Invoice on the allocation date", "Sales > Invoices > + New. Customer: the subscriber. Project: the estate. Date: the allocation date, not the date the money arrived. Item: Land sale. Confirm the line account is 4020, not Sales. Amount: the contract, not the cash. Tax: none. A development levy is a second line, 4042. A documentation fee is 4044. A corner piece is 4046. Save and mark it Open. A draft does not recognise the sale.")
    doc.figure(
        "invoice-new.png",
        "The top of a new invoice. Customer is the subscriber. Invoice date is the allocation date, or the handover date. The sample may show a foreign currency and a tax. Yours is naira, and tax is none.",
    )
    doc.figure(
        "invoice-lines.png",
        "The lines. Choose the item, then look at Tax before you save. It must be none on residential land and housing. Ignore State Sales tax and PayPal in the sample. A levy, a documentation fee and a corner piece are their own lines.",
    )
    doc.step("7", "Apply the retainer, then release the cost", "On the invoice, apply the retainer for this plot only. Type the unused amount, not a round figure and not another subscriber’s deposit. Then Accountant > Manual Journals > + New Journal. Debit 5020, credit 1210, the plot cost. Notes: Land allocation, the estate, the plot, the subscriber. Reporting method: Accrual and Cash. Publish. Save as Draft does not post.")
    doc.figure(
        "apply-retainer.png",
        "Apply the retainer to the invoice. Type only the unused amount for this plot. Do not apply a co-ownership balance, and do not apply another subscriber’s deposit. The sample amounts are not Baay’s.",
    )
    doc.figure(
        "journal.png",
        "A published journal. This is the shape of the cost release, the depreciation, the tax and the co-ownership accrual. In the sample the accounts are not Baay’s. Yours must balance, must say Published, and must name the plot in the notes.",
    )
    doc.p("Housing is the same shape, on a different day. The invoice is Housing sale, 4025, dated the handover. The cost is debit 5025, credit 1230. While the house is being built, cost sits in 1220. A receipt does not move it. GVC Gbagada and Baay Foreshore receipts are large. Size is not a handover.")
    doc.h2("Co-ownership, a refund, and a default charge")
    doc.p("A co-owner is not a subscriber. Do not raise a retainer. Categorise the receipt to 2025, on the project named in the certificate, or publish a journal debiting the bank and crediting 2025. The return accrues by journal: debit 7020, credit 2060, for the amount on that certificate, not a rate from this guide. Pay both parts only when the money leaves the bank. An issued cheque has not paid anything. A rollover stays in 2025. It is not a new receipt and it is not income.")
    doc.p("A deposit that goes back is a refund of the retainer payment. It debits 2020. It is not an expense, and it is not 6085. A sale that was already recognised is a credit note against that invoice, account 4020 or 4025, tax none, then a refund of the credit if the cash is going back.")
    doc.figure(
        "credit-note.png",
        "The credit note. Link it to the invoice. The account follows the original sale, 4020 or 4025, and tax is none. A deposit that was never a sale does not get a credit note. It gets a refund of the retainer payment.",
    )
    doc.p("The default charge is 10 per cent of an instalment that was actually missed. Finance raises it, not Client Relations, on an invoice, item Default charge, account 4120, tax none. It is not added to the retainer. It is not a tax. Do not let Zoho raise it automatically.")
    doc.h2("Reconcile, then lock")
    doc.step("8", "Lock only a month that is finished", "Uncategorized on each bank is nil, or every remaining line is already in 2045 or 1180 with a name and a date. The balance agrees to the statement. Petty cash and the site imprest agree to a count. Then Accountant > Transaction Locking. Lock Sales, Purchases, Banking and Accountant. The date is the last day of the month. The reason names the month, the person, and the statement date.")
    doc.figure(
        "lock-list.png",
        "Lock Sales, Purchases, Banking and Accountant. Do not lock the bank and leave journals open, or a late journal can still move a locked bank.",
    )
    doc.figure(
        "lock-date.png",
        "The lock date is the last day of the month you have reconciled. Write the reason. Do not leave the sample sentence. A locked month stays locked unless the Head of Accounts lifts it, and puts it back the same day.",
    )
    doc.p("The accountant does not unlock. The finance manager does not unlock without the Head of Accounts. A correction after the lock is a journal in the open month, with a note that says which month it belongs to.")
    doc.h2("The practice")
    doc.bullets([
        "Find the ₦1,000 retainer Client Relations left on Test Subscriber. Confirm the credit is 2020. Void the payment. Void the retainer. Confirm the bank and 2020 are back where they started.",
        "Publish a journal for ₦1, debit 1180, credit 2045. Reverse it. Both accounts must be unchanged.",
        "Point at the Categorize button on a deposit and say which account you would refuse.",
    ])
    doc.stop("Do not post the ₦3,000,000 example, the ₦45,000,000 handover, or any real subscriber, in this sitting. If the void does not clear the ₦1,000, do not leave the room. A practice receipt left in the bank will be matched to a real statement later.")
    doc.table(
        ["Finance may", "Finance may not"],
        [
            ["Match a bank line to a retainer payment or a bill payment", "Accept Other Income on a subscriber deposit"],
            ["Raise the invoice on allocation or handover, and the levy, fee and corner piece", "Recognise housing because the receipt was large, or because the status says Completed"],
            ["Publish the cost journal the same day, and the co-ownership accrual from the certificate", "Accrue a return the certificate does not state, or debit 2025 for the return"],
            ["Lock a reconciled month", "Change the chart, or unlock a month without the Head of Accounts"],
        ],
        [8.6, 8.6],
    )

    doc.h1("5.  Head of Accounts")
    doc.p("This sitting is short, and it is last, so you can show a month that Finance has just reconciled rather than a theory. The Head of Accounts can do everything. The discipline is that they do not re-key a receipt Client Relations has already recorded, and they do not post a sale from a status.")
    doc.say("These five screens are the close. If they agree, the month can be locked. If they do not, the lock waits. A lock on a mess freezes the mess.")
    doc.table(
        ["Screen", "Where", "What you ask them to say"],
        [
            ["Cash", "1010, 1020, 1030, 1040, 1050, 1055", "Each balance agrees to the statement or the count."],
            ["Collections", "Payments Received, Retainer Payment, this month", "Every line has a project. The total is what Client Relations says was collected."],
            ["Not yet a sale", "Unearned Revenue, 2020", "It moved only because a retainer was applied, or refunded."],
            ["Co-ownership", "2025 beside 2060", "Principal is in 2025. Unpaid return is in 2060. Neither is in 4020."],
            ["Still unnamed", "2045 and 1180", "Each line has an owner and an age. A line older than the week is named or explained."],
        ],
        [3.2, 6.2, 7.8],
    )
    doc.figure(
        "collections-report.png",
        "Payments Received, customized. Choose Retainer Payment, not All. Save it, so next month starts here. This is the second of the five screens. The amounts in the picture are Zoho’s sample.",
    )
    doc.p("Then look at profit only long enough to test the rule. 4020 is the plots allocated this month, at contract price, not the bank. 4025 is the units handed over. 4042, 4044 and 4046 can each be named. 4099 Sales is nil. If 4020 equals the bank, the month was posted the old way. Do not lock it. Find the retainer that was categorised as income, reverse it, and post it again as a retainer.")
    doc.p("An unlock needs a reason, and the lock goes back the same day. Show the lock screen, already used in the Finance sitting, and have the Head of Accounts refuse a request that has no reason. That is the practice. Do not lift a real lock to demonstrate it.")
    doc.p("The chart is not a daily tool. If a signed opening line has no account, stop and add it in the chart file. Do not drop it into General Income to make a screen balance. Opening Balance Offset and Opening Balance Adjustments stay nil.")

    doc.h1("6.  If someone clicks the wrong button")
    doc.p("Fix it before the next sitting, and before the person leaves the chair. A wrong document that sits overnight gets matched to a real bank line.")
    doc.table(
        ["What they did", "What you do", "What you do not do"],
        [
            ["Created a second customer", "Mark the empty one inactive, after you have checked it has no documents", "Merge it by retyping the balance"],
            ["Raised an invoice instead of a retainer", "Void the invoice if it is still unpaid and was never a sale", "Apply a receipt to it and call that a correction"],
            ["Categorised a deposit as Other Income", "Reverse that categorisation. Record the retainer. Match the line.", "Leave it, and journal the difference at month-end"],
            ["Recorded a customer payment", "Delete or void the unused credit. Record the retainer payment.", "Apply the unused credit to a later invoice"],
            ["Posted the ₦1,000 test and left it", "Void the payment and the retainer. Confirm 2020.", "Match it to a real statement so the bank will reconcile"],
            ["Published a journal to the wrong account", "Reverse the journal. Publish the right one. Do not edit a locked month.", "Unlock the month to tidy a draft"],
        ],
        [4.6, 6.4, 6.2],
    )
    doc.p("The posting guide has the account for each of these. Follow that note for the account, and follow this note for the order of the click. If a button has moved, search Settings for the document name. The account does not move because a menu has been renamed.")

    doc.h1("7.  Close the day")
    doc.p("Each person signs the line for the sitting they attended. A signature means they did the practice, not that they listened. You sign the last line only when the test receipt is voided and 2020 is back to its opening figure.")
    doc.table(
        ["Name", "Sitting", "Practice done", "Signature"],
        [
            ["", "Sales — found a subscriber and did not create one", "", ""],
            ["", "Client Relations — retainer of ₦1,000 on Head Office, account 2020", "", ""],
            ["", "Finance — voided the ₦1,000, and reversed the ₦1 journal", "", ""],
            ["", "Head of Accounts — read the five screens, and refused an unlock with no reason", "", ""],
        ],
        [4.0, 8.2, 2.4, 2.6],
    )
    doc.p("Consultant: ______________________________    Date: ____________________")
    doc.p("Head of Accounts: ________________________    Date: ____________________")
    doc.p("Give Sales their card, and Client Relations their card. Give Finance the posting guide. This training guide stays with you and with the Head of Accounts. It is not a tax ruling, and it is not a reload of a year that is already closed.")
    doc.p("Prepared for the consultant training Baay Projects Limited. " + DATE + ". The account codes are the chart of the same date. The posting steps, where this note does not repeat them, are in the posting guide.")


def main():
    needed = [
        "customers-list.png", "customer-menu.png", "payments-report.png",
        "new-customer.png", "project-menu.png", "retainer-new.png",
        "retainer-payment.png", "bank-match.png", "bank-categorize.png",
        "bill-account.png", "vendor-payment.png", "record-expense.png",
        "invoice-new.png", "invoice-lines.png", "apply-retainer.png",
        "journal.png", "credit-note.png", "lock-list.png", "lock-date.png",
        "collections-report.png",
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
