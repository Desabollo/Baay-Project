"""Consultant setup guide for Baay Projects Limited on Zoho Books Premium.

This is the configuration manual. It is not the posting guide. Every step is
something the consultant does once, before the first live receipt is posted.
The posting rules live in the posting guide and are only restated here where
a setting would otherwise post the wrong way.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from reportlab.lib.colors import Color, white
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    CondPageBreak,
    KeepTogether,
    ListFlowable,
    ListItem,
    Paragraph,
    Preformatted,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Setup_Guide.docx"
PDF = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Setup_Guide.pdf"

NAVY = RGBColor(0x1F, 0x4E, 0x79)
TEAL = RGBColor(0x1F, 0x6F, 0x78)
INK = RGBColor(0x1C, 0x28, 0x30)
MUTED = RGBColor(0x5E, 0x6A, 0x72)

PN = Color(0x1F / 255, 0x4E / 255, 0x79 / 255)
PT = Color(0x1F / 255, 0x6F / 255, 0x78 / 255)
PI = Color(0x1C / 255, 0x28 / 255, 0x30 / 255)
PM = Color(0x5E / 255, 0x6A / 255, 0x72 / 255)
RULE = Color(0xD0 / 255, 0xD7 / 255, 0xE2 / 255)
PALE = Color(0xF4 / 255, 0xF7 / 255, 0xF8 / 255)
CREAM = Color(0xF4 / 255, 0xE8 / 255, 0xD4 / 255)
RED = Color(0x8C / 255, 0x4A / 255, 0x3A / 255)


def shade(paragraph, hex_colour):
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_colour)
    p_pr.append(shd)


def set_run(run, name, size, bold=False, colour=INK, italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = colour


def add_text(paragraph, text, name="Calibri", size=11, bold=False, colour=INK, italic=False):
    run = paragraph.add_run(text)
    set_run(run, name, size, bold, colour, italic)
    return run


def para(doc, text, size=11, bold=False, colour=INK, space_before=2, space_after=6, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.12
    add_text(p, text, size=size, bold=bold, colour=colour, italic=italic)
    return p


def h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    add_text(p, text, "Calibri", 16, True, NAVY)
    return p


def h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    add_text(p, text, "Calibri", 13, True, TEAL)
    return p


def bullet(doc, text, bold_lead=None):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.75)
    p.paragraph_format.first_line_indent = Cm(-0.35)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.08
    add_text(p, "•  ", size=11, colour=TEAL)
    if bold_lead:
        add_text(p, bold_lead, size=11, bold=True)
        add_text(p, text, size=11)
    else:
        add_text(p, text, size=11)
    return p


def step(doc, number, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.85)
    p.paragraph_format.first_line_indent = Cm(-0.55)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.08
    add_text(p, f"{number}.  ", size=11, bold=True, colour=NAVY)
    add_text(p, text, size=11)
    return p


def callout(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.left_indent = Cm(0.3)
    p.paragraph_format.right_indent = Cm(0.2)
    shade(p, "F4E8D4")
    add_text(p, text, size=10.5, italic=True, colour=RGBColor(0x5A, 0x3E, 0x1B))
    return p


def table(doc, headers, rows, widths):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    t.autofit = False
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        add_text(p, h, size=9, bold=True, colour=RGBColor(0xFF, 0xFF, 0xFF))
        shade(p, "1F4E79")
        cell.width = Cm(widths[i])
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = t.rows[r + 1].cells[c]
            cell.text = ""
            p = cell.paragraphs[0]
            add_text(p, value, size=8.5)
            if r % 2:
                shade(p, "F4F7F8")
            cell.width = Cm(widths[c])
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t


def header_footer(section):
    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_text(p, "Baay Projects Limited", size=9, bold=True, colour=NAVY)
    add_text(p, "   ·   Zoho Books Premium setup guide", size=9, colour=MUTED)
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_text(fp, "Consultant configuration  ·  not the posting guide", size=8, colour=MUTED)
    add_text(fp, "          ", size=8)
    run = fp.add_run()
    set_run(run, "Calibri", 8, colour=MUTED)
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


def build_docx():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.7)
    section.right_margin = Cm(1.7)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.7)
    section.header_distance = Cm(0.6)
    section.footer_distance = Cm(0.5)
    header_footer(section)
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.font.color.rgb = INK

    para(doc, "BAAY PROJECTS LIMITED", 12, True, TEAL, 0, 0)
    para(doc, "Zoho Books Premium", 22, True, NAVY, 0, 0)
    para(doc, "Setup guide for the consultant", 16, True, TEAL, 0, 4)
    para(
        doc,
        "Every configuration, in the order it has to be done, before the first live receipt is posted. Prepared from the controlled chart and from the registers the accountant placed in the BAAY Documents folder in November 2026.",
        11, False, INK, 2, 6,
    )
    callout(
        doc,
        "This guide configures the organisation. It does not teach a person how to post a receipt, a bill or a journal. That is the posting guide, and it uses the account codes this setup puts in place. Do not start posting until the sign-off at the end of this guide is ticked.",
    )

    h1(doc, "1.  What you are configuring, and what you are not")
    para(doc, "Baay Projects Limited sells land and houses, builds for others, and takes co-ownership funds into named projects. The registers separate those streams. The software has to separate them too, or the first month will not be auditable.")
    bullet(doc, "A subscriber receipt is not income. It sits as a liability until a plot is allocated or a house is handed over.", "Receipts. ")
    bullet(doc, "Co-ownership is not a sale. The contract returns principal plus a return at the end of a period. Principal is a liability. The return is a finance cost.", "Co-ownership. ")
    bullet(doc, "A movement between the company's own banks is a transfer. It is never income and never an expense.", "Transfers. ")
    bullet(doc, "Estates are projects, not accounts. A new phase is a new project. Do not add an account because a new estate has opened.", "Estates. ")
    bullet(doc, "Residential land and housing are not charged to VAT. The organisation is still VAT-enabled, because a non-residential job may need it. The default on a land or housing item is no tax.", "VAT. ")
    bullet(doc, "The receipt states a default charge of 10 per cent of a missed instalment. That is Late Fee Income, code 4120. It is not a tax.", "Default charge. ")
    para(doc, "Do not type a subscriber's phone number, email address, date of birth or bank account number into any setting, template or custom field default. Those live in the customer record, and only as far as Zoho needs them to collect.")

    h1(doc, "2.  Before you open Zoho")
    para(doc, "Have these in hand. Do not start the organisation and then go looking for them.")
    table(
        doc,
        ["Item", "Where it comes from", "Why you need it before setup"],
        [
            ["Premium subscription", "The company, not the consultant", "Projects, retainer invoices and transaction locking are Premium features. A lower plan will not run this design."],
            ["Organisation admin", "A director, or the Head of Accounts", "Only an admin can import the chart, create roles and turn on locking."],
            ["Chart CSV", "BAAY_PROJECTS_LIMITED_Zoho_Books_Chart_of_Accounts.csv", "136 accounts. Import this file. Do not retype it."],
            ["Signed position at 31 December 2025", "The signed financial statements", "Opening balances come from that statement. This guide does not re-key them."],
            ["Bank statements, four banks", "Providus, First Bank, Sterling, GTBank", "Needed to connect feeds or to set the import format. Not needed to create the accounts."],
            ["Project list in section 8", "The November 2026 registers", "Create the projects before the first customer, so a receipt can be tagged on day one."],
            ["Logo and letterhead", "The company", "Receipts and retainer invoices go to subscribers. They should look like Baay."],
        ],
        [4.2, 6.2, 6.8],
    )
    callout(doc, "Do not import the Review sheet of the workbook, and do not import the sample chart that shipped with an earlier file. Account IDs in that sample were damaged by Excel. This CSV has a blank Account ID column on purpose.")

    h1(doc, "3.  Create the organisation")
    step(doc, 1, "Sign in to Zoho Books with the admin address the company will keep. Do not use a personal address that will leave with a staff member.")
    step(doc, 2, "Create an organisation named Baay Projects Limited. The legal name on the invoice is the company name, not a trading name, unless a director instructs otherwise in writing.")
    step(doc, 3, "Country or region: Nigeria. This sets the currency, the tax region and the address format. Do not pick another country and then try to correct it. Several tax settings cannot be changed afterwards.")
    step(doc, 4, "Base currency: NGN, Nigerian Naira. Symbol ₦. Decimal places: 2. This cannot be changed once transactions exist, so confirm it before the chart is imported.")
    step(doc, 5, "Fiscal year: January to December. The signed statements are 31 December. Do not set a March year.")
    step(doc, 6, "Language: English. Time zone: Africa/Lagos. Date format: dd MMM yyyy, so a receipt reads 3 Nov 2026 and cannot be mistaken for an American date.")
    step(doc, 7, "Industry, if asked: Real estate, or Construction. The choice does not change the chart. Do not accept a sample chart the wizard offers. You will import one.")
    step(doc, 8, "Plan: Zoho Books Premium. Confirm the plan on the subscription page before you continue. If the organisation is on a trial, note the trial end date in the sign-off and do not go live on the last day of the trial.")

    h1(doc, "4.  Organisation profile")
    para(doc, "Settings, then Organisation Profile. Complete every field. A blank address prints on the subscriber's receipt.")
    bullet(doc, "Baay Projects Limited.", "Organisation name. ")
    bullet(doc, "The registered office the company currently uses. The co-ownership contract template prints Plot 63, CMD Road, Magodo Phase 2, Lagos State. Confirm with the Head of Accounts before you save, and do not invent a second address.", "Address. ")
    bullet(doc, "The company's phone and email, not a consultant's and not a salesperson's personal line.", "Phone and email. ")
    bullet(doc, "Upload the company logo. Keep it small enough that the retainer invoice still fits the amount on the first screen.", "Logo. ")
    bullet(doc, "Leave the website blank if you have not confirmed it. Do not guess.", "Website. ")
    para(doc, "Save. Then open a sample retainer invoice in print preview and read the address aloud. If it is wrong, fix it now. Subscribers will keep these receipts.")

    h1(doc, "5.  Import the chart of accounts")
    para(doc, "Do this before items, before tax defaults, and before any opening balance. Several later settings point at an account, and the account has to exist first.")
    step(doc, 1, "Go to Accountant, Chart of Accounts.")
    step(doc, 2, "Choose Import, then Chart of Accounts. Upload BAAY_PROJECTS_LIMITED_Zoho_Books_Chart_of_Accounts.csv. Encoding UTF-8. The file is under 1 MB, inside Zoho's 10 MB limit.")
    step(doc, 3, "Duplicate handling: Overwrite. Zoho matches an existing account on the name. Overwrite is what writes the four-digit code onto the accounts already in the organisation. Skip would leave those accounts uncoded.")
    step(doc, 4, "Map Account Name, Account Code, Description, Account Type, Parent Account, Currency and Account Status. Do not map Account ID. Leave Mileage Rate, Mileage Unit and IsMileage on their defaults. They are unused.")
    step(doc, 5, "Currency on every row is NGN. Status on every row is Active. If the mapper asks for a currency and the file column is blank on a row, stop and check you uploaded the right file.")
    step(doc, 6, "Import. Parents are listed before their children. A child has the same account type as its parent. Both are Zoho rules, and the file already obeys them.")
    step(doc, 7, "Read the import result. A row that fails is almost always a locked account that was given a parent. This file leaves the parent blank on those accounts. If one still fails, skip that one row, do not edit the CSV to force a parent, and note it on the sign-off.")
    para(doc, "These accounts are coded, marked red in the workbook, and have no parent in the upload because Zoho locks them. Leave them where Zoho holds them. Do not post estate transactions to Sales, Cost of Goods Sold or Inventory Asset.")
    table(
        doc,
        ["Code", "Locked account", "What it is for"],
        [
            ["1090", "Undeposited Funds", "A receipt recorded before it reaches the bank. Clear it the same day."],
            ["1100", "Accounts Receivable", "The only customer-invoice control. Do not create a second one."],
            ["2000", "Accounts Payable", "The only supplier-bill control. Do not create a second one."],
            ["2150", "VAT Payable", "Output less input. No VAT on residential land or housing."],
            ["3020", "Retained Earnings", "Profit closes here. Do not journal the year into it by hand."],
            ["3098", "Opening Balance Offset", "Must be nil once opening balances tie."],
            ["2020", "Unearned Revenue", "Subscriber receipts, and the retainer liability. Set this before the first receipt."],
            ["4095", "Discount", "Invoice discounts only."],
            ["4120", "Late Fee Income", "The 10 per cent default charge. Not a tax."],
            ["4099", "Sales", "Do not use. Point items at 4010 to 4046."],
            ["5090", "Cost of Goods Sold", "Inventory module only. Plot and house releases go to 5020 and 5025."],
            ["1290", "Inventory Asset", "Inventory module only. Land and houses go to 1210, 1220 and 1230."],
        ],
        [2.2, 5.4, 9.6],
    )
    para(doc, "After the import, open Chart of Accounts and confirm the count is 136 active accounts. Then confirm these six codes by eye, because a wrong one will mis-state the first month: 2020 Unearned Revenue, 2025 Co-ownership liabilities, 2060 Co-ownership return payable, 4042 Revenue - development levy, 4120 Late Fee Income, 7020 Co-ownership return.")
    callout(doc, "A header account is not a posting account. Once it has children, Zoho will refuse a transaction on it. That is what you want. Do not turn a header into a posting account to make an import warning go away.")

    h1(doc, "6.  Taxes")
    para(doc, "Settings, Taxes. Nigeria. The organisation records VAT where a supply is actually subject to it. Residential land and housing are not. Do not let a default tax code put VAT on a plot.")
    step(doc, 1, "Confirm the VAT rate Zoho offers for Nigeria. Do not invent a second rate.")
    step(doc, 2, "Create a tax exemption, or a zero rate if the edition will not save an exemption, named Residential property — no VAT. The name should tell the next person why the line is untaxed.")
    step(doc, 3, "Set that exemption as the default on every land item and every housing item you create in the next section. A construction item that is genuinely subject to VAT gets the standard rate, and only that item.")
    step(doc, 4, "VAT Payable must remain account 2150. If a tax code wants to create its own liability account, point it at 2150. Do not allow a second VAT liability.")
    step(doc, 5, "Do not enable a tax on retainers. A subscriber receipt is not a supply. Tax, if any, belongs on the invoice that recognises the sale, and on these sales that tax is none.")
    step(doc, 6, "Withholding tax is not a tax code on the sales invoice. When a customer deducts WHT, the receipt is posted as the net bank amount plus account 1150. Set nothing in Taxes for that. The posting guide shows the entry.")
    callout(doc, "This setup is not a tax opinion. If a supply is not residential land or housing, ask the Head of Accounts before you leave the standard rate off. Do not decide a new tax treatment inside a customer record.")

    h1(doc, "7.  Items")
    para(doc, "Items are the products. Projects are the estates. A line on a retainer or an invoice is an item, tagged to a project. Create the items once. Do not create an item per estate, and do not create an item per subscriber.")
    para(doc, "Items, then New Item. Each item below is a service, not goods, so Zoho does not try to move Inventory Asset. Sellable: yes. Purchasable: no, except the two cost items, which are purchasable and not sellable.")
    table(
        doc,
        ["Item name", "Sales account", "Tax", "Used on"],
        [
            ["Land subscription", "2020 Unearned Revenue", "No VAT", "Retainer invoice for a plot"],
            ["Housing subscription", "2020 Unearned Revenue", "No VAT", "Retainer invoice for a house or unit"],
            ["Land sale", "4020 Revenue - land and property sales", "No VAT", "Invoice raised on allocation"],
            ["Housing sale", "4025 Revenue - housing", "No VAT", "Invoice raised on handover"],
            ["Construction contract", "4010 Revenue - civil engineering and construction", "Ask before taxing", "Certified construction invoice"],
            ["Project management", "4030 Revenue - project management and consulting", "Ask before taxing", "Consulting invoice"],
            ["Development levy", "4042 Revenue - development levy", "No VAT", "Invoice, on allocation, if the register shows a levy"],
            ["Documentation fee", "4044 Revenue - documentation fee", "No VAT", "Invoice, on allocation, if a documentation fee was charged"],
            ["Corner piece", "4046 Revenue - corner piece", "No VAT", "Invoice, on allocation, if the plot is a corner piece"],
            ["Default charge", "4120 Late Fee Income", "No VAT", "Only after an instalment is actually missed"],
            ["Co-ownership principal", "2025 Co-ownership liabilities", "No tax", "Retainer-style receipt. Never an income item."],
        ],
        [4.4, 6.6, 3.2, 2.8],
    )
    para(doc, "The last column is squeezed on purpose. Read it with the posting guide before you raise the first document. Two further settings on every item:")
    bullet(doc, "Leave the rate blank, or at zero. Prices differ by plot and by house. A default rate will be raised in error.", "Rate. ")
    bullet(doc, "Description: one line, for example “Plot subscription, allocated later”. Do not put an estate name in the item description. The project carries the estate.", "Description. ")
    callout(doc, "If Zoho will not save a sales account of type Other Current Liability on an item, do not point the item at an income account to get past the error. Raise the subscriber document as a retainer invoice, which posts to 2020, and keep the item list for the invoice that is raised on allocation or handover. Note the limitation on the sign-off.")

    h1(doc, "8.  Projects")
    para(doc, "Time Tracking, Projects, New Project. One project per estate or phase. The customer on the project may be left blank where Zoho allows it. Where it does not, do not attach the estate to the first subscriber who happens to pay. Create the project against the company, or leave the field until the form accepts a blank, and record which you did.")
    para(doc, "Billing method: fixed cost of the project, or based on task hours, whichever the form requires. Do not turn on recurring billing. Baay does not bill a subscriber every month by itself. Currency: NGN. Watchlist: on, so a quiet project is still visible.")
    table(
        doc,
        ["Project name", "Type", "Where the registers name it"],
        [
            ["Green City Phase 1 (Ketu-Epe)", "Land", "Client register, 2026"],
            ["Green City Phase 2 (Ketu-Epe)", "Land", "Client register, 2026"],
            ["Green City Phase 3 (Omu-Epe)", "Land", "Client register, 2026"],
            ["Green City Phase 3 Extension (Omu-Epe)", "Land", "Client register, 2026"],
            ["Green City Ibadan (Akinyele-Moniya)", "Land", "Client register, 2026"],
            ["Heritage Garden", "Land", "Client register and the co-ownership contract"],
            ["Green Hillside", "Land", "Client register, 2026"],
            ["Pacific Apartment", "Housing", "Client register and the Pacific expense file"],
            ["Pacific Apartment B", "Housing", "Revenue and expense register"],
            ["Eden Residence", "Housing", "Eden revenue and expense file"],
            ["GVC Gbagada", "Housing", "Client register, 2026"],
            ["Baay Foreshore", "Housing", "Client register and the sample receipt"],
            ["Baay Foreshore Ibadan", "Housing", "Revenue and expense register"],
            ["Ologolo Land", "Land", "Revenue and expense register"],
            ["12 Basket", "Construction", "Revenue register"],
            ["Scrapays", "Construction", "Revenue and expense register"],
            ["Engagement Service", "Service", "Revenue and expense register"],
            ["Agbowa", "Land", "Earlier project list. Confirm it is still live before you post to it."],
            ["Pacific Court", "Housing", "Earlier project list. Confirm it is still live before you post to it."],
            ["Gorge View Court", "Housing", "Earlier project list. The client register names GVC Gbagada separately. Do not merge them."],
        ],
        [7.4, 2.6, 7.2],
    )
    para(doc, "Also create two internal projects, so a cost that is not a sale still has a home: Head Office, and Ibadan Office. Staff-lodge costs at a site take the site's project, not a new one, and they take account 6065.")
    para(doc, "A corner piece is not a project. It is item Corner piece on the estate's project. A co-ownership contract names the project the funds go into. The receipt takes that project, and account 2025, not a project called Co-ownership.")
    callout(doc, "Do not delete a project because it has no transactions yet. An empty project is cheaper than a receipt posted to the wrong estate. Add a project the day a new phase is named. Do not wait for the first receipt.")

    h1(doc, "9.  Customers, vendors, and what not to import")
    para(doc, "Do not import the client workbook on day one. It holds phone numbers, email addresses, dates of birth and, in the commission file, agents' bank account numbers. A bulk import of that file into a live organisation is a data leak waiting for a report.")
    step(doc, 1, "Settings, Customers, and turn on the fields the posting needs: customer type, and a custom field named Register status, with the choices Prospect, Paying, Completed, Refund, Void, Incentive.")
    step(doc, 2, "Add a custom field named Plot or unit, plain text. This is the block and plot, such as Block D7. It is not an account and not a project.")
    step(doc, 3, "Add a custom field named Payment plan, with the choices the register uses: Outright, 3 Months, 6 Months, 12 Months, Other.")
    step(doc, 4, "Currency on every customer: NGN. Payment terms: due on receipt. Do not set a 30-day term. A subscriber is not a trade debtor until allocation or handover.")
    step(doc, 5, "Create the first customers by hand, or from a cut-down file that contains only the name, the project and the plot. Leave phone and email blank unless Client Relations is going to use Zoho to write to that person.")
    step(doc, 6, "Vendors: create a vendor when the first bill arrives. Type: supplier, subcontractor, agent, or authority. Do not import the commission workbook. Agents' account numbers stay in the commission file, outside Zoho, until the Head of Accounts decides otherwise.")
    callout(doc, "An agent is a vendor, not a customer. Commission is account 6035. It is never deducted from the subscriber's receipt and never deducted from revenue.")

    h1(doc, "10.  Bank accounts and the feed")
    para(doc, "The four bank accounts already exist, because the chart created them. Do not create them again from the banking screen. Connecting a feed to a second account of the same name is the usual way a transfer gets posted twice.")
    table(
        doc,
        ["Code", "Account", "Role in the books"],
        [
            ["1010", "Providus Bank", "Operations, as the master workbook describes it."],
            ["1020", "First Bank of Nigeria", "Collections."],
            ["1030", "Sterling Bank", "Disbursements."],
            ["1040", "Guaranty Trust Bank", "The fourth company account."],
        ],
        [2.2, 5.2, 9.8],
    )
    step(doc, 1, "Banking, then the account, then Connect. Prefer a direct feed. Where the bank has no feed, set up a statement import and save the column mapping: date, narration, debit, credit, balance.")
    step(doc, 2, "Statement date format must match the bank, not the organisation. A swapped day and month will import, and it will be wrong. Import one week, read five lines, and only then import history.")
    step(doc, 3, "Do not turn on automatic categorisation. Zoho will offer to treat a subscriber deposit as income or as a customer payment. Both are wrong for Baay. Every line is matched or categorised by a person.")
    step(doc, 4, "Bank rules, if you create any, may only do one thing: suggest a match to an existing retainer payment or bill payment. A rule must not create income, must not create an expense, and must not create a customer payment.")
    step(doc, 5, "Transfers between 1010, 1020, 1030 and 1040 are recorded with the Transfer action. There is no transfer account. Do not create one.")
    para(doc, "Petty cash is 1050 and site imprest is 1055. They are cash accounts, not banks. Do not connect a feed to them. A reimbursement of the float credits the float. It does not credit income.")

    h1(doc, "11.  Opening balances")
    para(doc, "Open from the signed statement of financial position at 31 December 2025. This guide does not contain those figures and does not ask you to reconstruct them.")
    step(doc, 1, "Settings, Opening Balances. Date: 31 December 2025. Enter each signed line against the account in this chart that describes it. Cash is the four banks, not one bank.")
    step(doc, 2, "Subscriber money held at that date, not yet allocated and not yet handed over, is Unearned Revenue 2020. It is not revenue and not a receivable.")
    step(doc, 3, "Co-ownership principal held at that date is 2025. Do not park it in 2020. The two liabilities are different promises.")
    step(doc, 4, "Opening Balance Offset 3098 and Opening Balance Adjustments 2090 must both be nil when you finish. A balance in either account means the opening trial balance does not tie. Do not go live with a balance in either.")
    step(doc, 5, "Do not enter 2026 transactions as opening balances. They are the first month's work, and the posting guide covers them.")
    callout(doc, "If a signed line has no account in this chart, stop. Add the account in the chart file first, re-import that one account, and note why it was added. Do not drop the balance into General Income or Other Expenses to make the screen balance.")

    h1(doc, "12.  Document numbers and templates")
    para(doc, "Settings, Preferences, then the document you are numbering. Set the prefix and the next number before the first document is saved. A number that has been used cannot be reused cleanly.")
    table(
        doc,
        ["Document", "Prefix", "Next number", "Who may raise it"],
        [
            ["Retainer invoice", "RET-", "0001", "Client Relations or Finance"],
            ["Invoice", "INV-", "0001", "Finance only"],
            ["Credit note", "CN-", "0001", "Finance only"],
            ["Customer payment", "PMT-", "0001", "Client Relations or Finance"],
            ["Vendor bill", "BILL-", "0001", "Finance only"],
            ["Vendor payment", "VPMT-", "0001", "Finance only"],
            ["Manual journal", "JNL-", "0001", "Finance only. A draft is not a posting."],
            ["Project", "No prefix", "Name is the estate", "Consultant at setup, then Finance"],
        ],
        [4.0, 2.6, 4.4, 6.2],
    )
    para(doc, "On the retainer invoice template, show the project, the plot or unit, the amount received, the contract amount and the balance. That is what the company's own receipt shows, and the subscriber already expects it. Hide the shipping field, the adjustment field, and any tax summary that would print a VAT line of zero in a way that looks like a charge.")
    para(doc, "Terms on the retainer invoice, in the company's words and no stronger: the balance is payable on the agreed plan, and a defaulted instalment carries a charge of 10 per cent of the amount defaulted. Do not add a penalty the receipt does not state.")
    para(doc, "Do not print a co-owner's return rate on a subscriber's land receipt. They are different documents. The co-ownership certificate carries the rate. The land receipt does not.")

    h1(doc, "13.  Users, roles, and what each person may do")
    para(doc, "Settings, Users and Roles. Create the roles below. Do not share one login. A shared login makes a lock useless, because you cannot see who posted.")
    table(
        doc,
        ["Person", "May do", "May not do"],
        [
            ["Consultant", "Everything in this guide, until sign-off", "Post a live receipt after hand-over, unless asked in writing"],
            ["Head of Accounts", "Every posting, every lock, the chart", "Nothing is reserved above this role"],
            ["Finance manager", "Bills, journals, categorise the bank, lock a month", "Change the chart, or unlock a month without the Head of Accounts"],
            ["Accountant", "Bills, journals, bank, reports", "Unlock a month. Change a tax code."],
            ["Client Relations", "Open a subscriber, raise a retainer, record a payment, set the status", "Categorise a bank line, publish a journal, touch the chart"],
            ["Sales", "See their project's customers and the balance on a retainer", "Raise an invoice, record a payment, or see the commission file"],
            ["Auditor, later", "Read only, for the year they are examining", "Post, export the customer list by email, or change a lock"],
        ],
        [3.4, 6.6, 7.2],
    )
    para(doc, "Turn on two-factor authentication for every user before you send the invitation. Remove the consultant's admin access when the sign-off is signed, or leave it only if the company has asked for a support window, and date that window.")

    h1(doc, "14.  Preferences that change a posting")
    para(doc, "Settings, Preferences. Work down the list. The defaults Zoho ships will post a subscriber deposit as income if you leave them.")
    bullet(doc, "Off. A month is not closed by a sales order.", "Sales orders. ")
    bullet(doc, "Off for land and housing. A subscriber is not invoiced when they pay. The invoice is raised on allocation or handover, and the retainer is applied to it.", "Invoice on payment. ")
    bullet(doc, "On. Client Relations needs it for a receipt.", "Retainer invoices. ")
    bullet(doc, "2020 Unearned Revenue. Confirm this on the retainer settings page before the first receipt. If it points at Sales, every receipt becomes income.", "Retainer liability account. ")
    bullet(doc, "Off. Nothing in these registers is delivered by a carrier.", "Shipping charges. ")
    bullet(doc, "Off. A late fee is raised by a person, as the default charge, and only when an instalment was missed. An automatic fee will fire on a plan that was never agreed in Zoho.", "Automatic late fees. ")
    bullet(doc, "Allow them, and restrict them to Finance. A journal is how an allocation, a handover and a co-ownership accrual are posted.", "Manual journals. ")
    bullet(doc, "On, at the organisation level. The first lock is the month before go-live, so nobody back-posts into the opening position.", "Transaction locking. ")
    bullet(doc, "Do not enable. A subscriber receipt must not become a sale because a date passed.", "Revenue recognition schedules. ")
    bullet(doc, "NGN only, unless a foreign-currency bill has actually arrived. If one has, add that currency and leave exchange differences on 9010.", "Multi-currency. ")
    bullet(doc, "On, so a deleted receipt leaves a trail. A subscriber will ask.", "Audit trail. ")

    h1(doc, "15.  Custom fields")
    para(doc, "Create only the fields a report needs. Every extra field is a field someone will fill with a phone number.")
    table(
        doc,
        ["On", "Field", "Type", "Why it exists"],
        [
            ["Customer", "Register status", "Dropdown", "Prospect, Paying, Completed, Refund, Void, Incentive. Matches the client register."],
            ["Customer", "Plot or unit", "Text", "Block and plot. The project is the estate, not the plot."],
            ["Customer", "Payment plan", "Dropdown", "Outright, 3, 6, 12 months, or Other."],
            ["Retainer", "Contract amount", "Amount", "The price. The retainer itself is only what has been received."],
            ["Retainer", "Plan end date", "Date", "So a missed instalment can be seen without opening the register."],
            ["Project", "Property type", "Dropdown", "Land, Housing, Construction, Service, Internal."],
            ["Journal", "Reason", "Dropdown", "Allocation, Handover, Co-ownership accrual, Reclass, Other."],
        ],
        [2.6, 3.6, 2.4, 8.6],
    )
    para(doc, "Do not add a field for commission rate, agent bank, or co-owner return rate on the customer. Commission lives on the vendor bill. The return rate lives on the co-ownership certificate, and the accrual is a journal.")

    h1(doc, "16.  Transaction locking")
    para(doc, "Accountant, Transaction Locking. Lock everything up to and including 31 December 2025 once the opening balances tie. Reason: opening position agreed to the signed statement. After go-live, Finance locks each month when the bank is matched and the retainer report agrees to 2020. The posting guide says how. This section only turns the feature on and proves it works.")
    step(doc, 1, "Lock to 31 December 2025.")
    step(doc, 2, "Try to save a journal dated in 2025 with a test user. Zoho must refuse it.")
    step(doc, 3, "Unlock it, as admin, and confirm a reason is required. Then lock it again. If a lock can be lifted without a reason, the preference is not finished.")

    h1(doc, "17.  Reports to save before go-live")
    para(doc, "Save these as custom reports, with the names below, so the accountant does not rebuild them in the first week.")
    bullet(doc, "Unearned Revenue 2020, by project. This is the subscriber liability.", "Subscriber funds. ")
    bullet(doc, "Co-ownership liabilities 2025, and Co-ownership return payable 2060, side by side.", "Co-ownership. ")
    bullet(doc, "Retainer invoices, filtered to those with a balance, by project.", "Still paying. ")
    bullet(doc, "Account 1180 and account 2045. Both should be worked to nil every week. A balance means a bank line nobody has named.", "Unidentified. ")
    bullet(doc, "Selling commissions 6035 against bills still open.", "Commission. ")
    bullet(doc, "Trial balance. Opening Balance Offset and Opening Balance Adjustments must show nil.", "Tie-out. ")
    para(doc, "Schedule the trial balance and the subscriber-funds report to the Head of Accounts every Monday morning. Do not schedule a report that includes customer phone numbers or email addresses.")

    h1(doc, "18.  What the first week is allowed to post")
    para(doc, "Setup is finished only when a test receipt has been raised and voided. Do it in this order, with a customer named Test Subscriber, on project Head Office.")
    step(doc, 1, "Raise a retainer invoice for ₦1,000 on item Land subscription. Confirm the credit is 2020, not an income account.")
    step(doc, 2, "Record a payment of ₦1,000 to Providus. Confirm the bank increased and 2020 moved to the payment. Void the payment. Confirm the bank went back.")
    step(doc, 3, "Void the retainer. Confirm 2020 is back to its opening figure.")
    step(doc, 4, "Raise a journal dated today, Dr 1180 Cr 2045, ₦1. Publish it. Then reverse it. Confirm both accounts are unchanged. This proves journals post only when published, and that a reversal is available.")
    step(doc, 5, "Delete the test customer only if Zoho allows it after the void. If it does not, mark the customer inactive and leave the voided documents. Do not leave a live ₦1,000 receipt in the bank.")
    callout(doc, "Do not run this test on a real subscriber. Do not run it on a date in a locked period. If any of the five steps posts to Sales, stop and correct the item or the retainer account before anyone is trained.")

    h1(doc, "19.  Hand-over")
    para(doc, "Train in two rooms. Client Relations and Sales get the retainer, the payment and the status field, and nothing else. Finance gets the bank match, the allocation journal, the handover journal, the co-ownership accrual and the lock. Give each room the posting guide, not this one. This guide stays with the Head of Accounts and the consultant.")
    para(doc, "On the day you hand over, send the Head of Accounts the organisation link, the list of users, the chart count, and the sign-off below. Take the consultant off the admin role, or date the support window.")

    h1(doc, "20.  Sign-off")
    para(doc, "Tick these in the organisation, not on a copy of this page. A tick without the screen open is not a tick.")
    for line in [
        "Plan is Premium, and the trial end date is written here if the subscription has not been paid: ____________________",
        "Country Nigeria, currency NGN, year January to December, time zone Africa/Lagos.",
        "Address on a printed retainer invoice has been read aloud and agreed.",
        "Chart imported with Overwrite. Active accounts: 136. Account ID was not mapped.",
        "2020, 2025, 2060, 4042, 4120 and 7020 opened and read.",
        "No VAT on the land items, the housing items, or the retainer.",
        "Projects created from the list in section 8, including Head Office and Ibadan Office.",
        "Retainer liability account is 2020.",
        "Automatic categorisation is off. No bank rule creates income.",
        "Opening balances tie. 3098 and 2090 are nil. The period to 31 December 2025 is locked.",
        "Test receipt of ₦1,000 was raised, voided, and is not sitting in the bank.",
        "Users created. Two-factor authentication is on. The consultant's admin access ends on: ____________________",
    ]:
        bullet(doc, line)
    para(doc, "Consultant: ______________________________    Date: ____________________", 11, False, INK, 10, 2)
    para(doc, "Head of Accounts: ________________________    Date: ____________________", 11, False, INK, 2, 2)
    para(doc, "Director: ________________________________    Date: ____________________", 11, False, INK, 2, 6)
    para(
        doc,
        "Prepared for Baay Projects Limited. Account codes are the chart imported in section 5. Posting steps are in the posting guide of the same date. This setup is not a tax ruling and not an audit opinion.",
        9, False, MUTED, 8, 2, True,
    )
    doc.save(DOCX)


def P(text, style):
    return Paragraph(text, style)


def pdf_table(styles, headers, rows, widths):
    data = [[P(h, styles["th"]) for h in headers]]
    for row in rows:
        data.append([P(c, styles["td"]) for c in row])
    tbl = Table(data, colWidths=widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PN),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, PALE]),
        ("GRID", (0, 0), (-1, -1), 0.3, RULE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return tbl


def build_pdf():
    styles = {
        "kicker": ParagraphStyle("kicker", fontName="Times-Bold", fontSize=9, leading=12, textColor=PT, tracking=0.4),
        "title": ParagraphStyle("title", fontName="Times-Bold", fontSize=22, leading=26, textColor=PN, spaceAfter=2),
        "sub": ParagraphStyle("sub", fontName="Times-Bold", fontSize=13, leading=16, textColor=PT, spaceAfter=8),
        "body": ParagraphStyle("body", fontName="Times-Roman", fontSize=10, leading=13.4, textColor=PI, alignment=TA_JUSTIFY, spaceAfter=7),
        "h1": ParagraphStyle("h1", fontName="Times-Bold", fontSize=13, leading=16, textColor=PN, spaceBefore=12, spaceAfter=4),
        "h2": ParagraphStyle("h2", fontName="Times-Bold", fontSize=11, leading=14, textColor=PT, spaceBefore=8, spaceAfter=3),
        "bullet": ParagraphStyle("bullet", fontName="Times-Roman", fontSize=10, leading=13, textColor=PI, leftIndent=12, bulletIndent=0, spaceAfter=2),
        "small": ParagraphStyle("small", fontName="Times-Italic", fontSize=8.5, leading=11, textColor=PM, spaceBefore=4),
        "th": ParagraphStyle("th", fontName="Times-Bold", fontSize=8, leading=10, textColor=white),
        "td": ParagraphStyle("td", fontName="Times-Roman", fontSize=7.5, leading=9.5, textColor=PI),
        "call": ParagraphStyle("call", fontName="Times-Italic", fontSize=9, leading=12, textColor=Color(0x5A / 255, 0x3E / 255, 0x1B / 255)),
        "step": ParagraphStyle("step", fontName="Times-Roman", fontSize=10, leading=13, textColor=PI, leftIndent=14, firstLineIndent=-14, spaceAfter=2),
    }

    def on_page(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(PN)
        canvas.rect(0, A4[1] - 12 * mm, A4[0], 12 * mm, fill=1, stroke=0)
        canvas.setFillColor(white)
        canvas.setFont("Times-Bold", 8.5)
        canvas.drawString(16 * mm, A4[1] - 7.4 * mm, "Baay Projects Limited")
        canvas.setFont("Times-Roman", 8.5)
        canvas.drawRightString(A4[0] - 16 * mm, A4[1] - 7.4 * mm, "Zoho Books Premium  ·  Setup guide")
        canvas.setFillColor(PT)
        canvas.rect(0, 0, A4[0], 9 * mm, fill=1, stroke=0)
        canvas.setFillColor(white)
        canvas.setFont("Times-Roman", 8)
        canvas.drawString(16 * mm, 3.6 * mm, "Consultant configuration  ·  not the posting guide")
        canvas.drawRightString(A4[0] - 16 * mm, 3.6 * mm, str(doc.page))
        canvas.restoreState()

    story = []
    story.append(P("BAAY PROJECTS LIMITED", styles["kicker"]))
    story.append(Spacer(1, 2 * mm))
    story.append(P("Zoho Books Premium", styles["title"]))
    story.append(P("Setup guide for the consultant", styles["sub"]))
    story.append(P(
        "Every configuration, in the order it has to be done, before the first live receipt is posted. Prepared from the controlled chart and from the registers the accountant placed in the BAAY Documents folder in November 2026.",
        styles["body"],
    ))
    story.append(Table(
        [[P("This guide configures the organisation. It does not teach a person how to post a receipt, a bill or a journal. That is the posting guide. Do not start posting until the sign-off at the end of this guide is ticked.", styles["call"])]],
        colWidths=[178 * mm],
    ))
    story[-1].setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), CREAM),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    story.append(P("1.  What you are configuring, and what you are not", styles["h1"]))
    story.append(P("Baay Projects Limited sells land and houses, builds for others, and takes co-ownership funds into named projects. The registers separate those streams. The software has to separate them too, or the first month will not be auditable.", styles["body"]))
    for line in [
        "A subscriber receipt is not income. It sits as a liability until a plot is allocated or a house is handed over.",
        "Co-ownership is not a sale. The contract returns principal plus a return. Principal is a liability. The return is a finance cost.",
        "A movement between the company's own banks is a transfer.",
        "Estates are projects, not accounts. A new phase is a new project.",
        "Residential land and housing are not charged to VAT.",
        "A missed instalment carries a default charge of 10 per cent. That is Late Fee Income, 4120, not a tax.",
    ]:
        story.append(P("•  " + line, styles["bullet"]))
    story.append(Spacer(1, 2 * mm))
    story.append(P("Do not type a subscriber's phone number, email address, date of birth or bank account number into a setting, a template or a custom-field default.", styles["body"]))

    story.append(P("2.  Before you open Zoho", styles["h1"]))
    story.append(P("Have these in hand. Do not start the organisation and then go looking for them.", styles["body"]))
    story.append(pdf_table(styles, ["Item", "Why you need it before setup"], [
        ["Premium subscription", "Projects, retainer invoices and transaction locking are not on a lower plan."],
        ["An admin login the company will keep", "Only an admin can import the chart, create roles and turn on locking."],
        ["The chart CSV", "136 accounts. Import the file. Do not retype it and do not import the Review sheet."],
        ["Signed position at 31 December 2025", "Opening balances come from that statement. This guide does not re-key them."],
        ["The project list in section 8", "Create the projects before the first customer, so a receipt can be tagged on day one."],
        ["Logo and the confirmed office address", "Receipts go to subscribers. Confirm the address. Do not invent one."],
    ], [62 * mm, 116 * mm]))
    story.append(Spacer(1, 2 * mm))
    story.append(P("Do not import an earlier sample chart. Account IDs in that sample were damaged by Excel. This CSV leaves Account ID blank on purpose.", styles["body"]))

    story.append(P("3.  Create the organisation", styles["h1"]))
    for n, line in enumerate([
        "Sign in with the admin address the company will keep. Do not use a personal address.",
        "Name the organisation Baay Projects Limited.",
        "Country: Nigeria. Do not pick another country and correct it later. Several tax settings will not change.",
        "Base currency: NGN, two decimal places. This cannot be changed once transactions exist.",
        "Fiscal year: January to December. Time zone: Africa/Lagos. Date format: dd MMM yyyy.",
        "Decline any sample chart the wizard offers. Plan: Zoho Books Premium. Write down the trial end date if the subscription is not yet paid.",
    ], 1):
        story.append(P(f"{n}.  {line}", styles["step"]))

    story.append(P("4.  Organisation profile", styles["h1"]))
    story.append(P("Settings, Organisation Profile. Name, the registered office the Head of Accounts confirms, the company's phone and email, and the logo. The co-ownership contract template prints Plot 63, CMD Road, Magodo Phase 2, Lagos State. Confirm that address before you save it. Do not invent a second one. Print a sample retainer invoice and read the address aloud before you leave this screen.", styles["body"]))

    story.append(P("5.  Import the chart of accounts", styles["h1"]))
    for n, line in enumerate([
        "Accountant, Chart of Accounts, Import, Chart of Accounts.",
        "Upload the CSV. Encoding UTF-8. The file is under 1 MB, inside Zoho's 10 MB limit.",
        "Duplicate handling: Overwrite. Zoho matches on the name. Overwrite writes the four-digit code onto accounts already in the organisation. Skip would leave them uncoded.",
        "Map Account Name, Account Code, Description, Account Type, Parent Account, Currency and Account Status. Do not map Account ID.",
        "Import. Parents are listed before their children, and a child has the same type as its parent.",
        "If a row fails, it is almost always a locked account. Skip that row. Do not edit the CSV to force a parent onto it.",
    ], 1):
        story.append(P(f"{n}.  {line}", styles["step"]))
    story.append(Spacer(1, 1.5 * mm))
    story.append(P("Locked, coded, and left without a parent: Undeposited Funds, Accounts Receivable, Accounts Payable, VAT Payable, Retained Earnings, Opening Balance Offset, Unearned Revenue, Discount, Late Fee Income, Shipping Charge, Other Charges, Opening Balance Adjustments, Dimension Adjustments, Inventory Asset, Cost of Goods Sold, Sales, and Exchange Gain or Loss. Do not post estate transactions to Sales, Cost of Goods Sold or Inventory Asset.", styles["body"]))
    story.append(P("Confirm 136 active accounts. Then open and read 2020 Unearned Revenue, 2025 Co-ownership liabilities, 2060 Co-ownership return payable, 4042 Revenue - development levy, 4120 Late Fee Income and 7020 Co-ownership return. A header that has children will refuse a transaction. Leave it that way.", styles["body"]))

    story.append(P("6.  Taxes", styles["h1"]))
    for n, line in enumerate([
        "Settings, Taxes. Confirm the Nigeria VAT rate Zoho offers. Do not invent a second rate.",
        "Create an exemption named Residential property — no VAT. Use a zero rate only if the edition will not save an exemption.",
        "Set that exemption on every land item and every housing item. A construction item takes the standard rate only if the Head of Accounts says the supply is taxable.",
        "Point every tax code at VAT Payable 2150. Do not allow a second VAT liability.",
        "Do not enable a tax on retainers. A subscriber receipt is not a supply.",
        "Do not build a withholding-tax code on the sales invoice. A customer who deducts it is posted as the net bank amount plus account 1150. The posting guide shows the entry.",
    ], 1):
        story.append(P(f"{n}.  {line}", styles["step"]))
    story.append(Spacer(1, 1.5 * mm))
    story.append(P("This setup is not a tax opinion. If a supply is not residential land or housing, ask before you leave the standard rate off. Do not decide a new treatment inside a customer record.", styles["body"]))

    story.append(P("7.  Items", styles["h1"]))
    story.append(P("Items are the products. Projects are the estates. Do not create an item per estate or per subscriber. Each item is a service. Rate blank. Tax none, except a construction item the Head of Accounts has said is taxable.", styles["body"]))

    item_header = ["Item", "Sales account", "Used when"]
    item_rows = [
        ["Land subscription", "2020 Unearned Revenue", "Retainer for a plot"],
        ["Housing subscription", "2020 Unearned Revenue", "Retainer for a house"],
        ["Land sale", "4020 Land and property", "Invoice on allocation"],
        ["Housing sale", "4025 Housing", "Invoice on handover"],
        ["Construction contract", "4010 Construction", "Certified invoice"],
        ["Project management", "4030 Consulting", "Consulting invoice"],
        ["Development levy", "4042 Development levy", "If the register shows a levy"],
        ["Documentation fee", "4044 Documentation fee", "If a documentation fee was charged"],
        ["Corner piece", "4046 Corner piece", "If the plot is a corner"],
        ["Default charge", "4120 Late Fee Income", "Only after a missed instalment"],
        ["Co-ownership principal", "2025 Co-ownership liabilities", "Never an income item"],
    ]
    data = [[P(h, styles["th"]) for h in item_header]]
    for row in item_rows:
        data.append([P(c, styles["td"]) for c in row])
    tbl = Table(data, colWidths=[52 * mm, 68 * mm, 58 * mm], repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PN),
        ("BACKGROUND", (0, 1), (-1, -1), white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, PALE]),
        ("GRID", (0, 0), (-1, -1), 0.3, RULE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 2 * mm))
    story.append(P("If Zoho will not save a liability account on an item, do not point the item at income to get past the error. Raise the subscriber document as a retainer invoice, which posts to 2020, and keep the items for the invoice on allocation or handover. Note the limitation on the sign-off.", styles["body"]))

    story.append(P("8.  Projects", styles["h1"]))
    story.append(P("Time Tracking, Projects, New Project. One project per estate or phase. Currency NGN. Do not turn on recurring billing. Do not attach an estate to the first subscriber if the customer field can be left blank. Create these, and also Head Office and Ibadan Office.", styles["body"]))
    projects = [
        "Green City Phase 1 (Ketu-Epe), Phase 2 (Ketu-Epe), Phase 3 (Omu-Epe), Phase 3 Extension (Omu-Epe), and Green City Ibadan (Akinyele-Moniya).",
        "Heritage Garden, Green Hillside, Ologolo Land, Agbowa.",
        "Pacific Apartment, Pacific Apartment B, Eden Residence, GVC Gbagada, Baay Foreshore, Baay Foreshore Ibadan, Pacific Court, Gorge View Court.",
        "12 Basket, Scrapays, and Engagement Service.",
    ]
    for line in projects:
        story.append(P("•  " + line, styles["bullet"]))
    story.append(Spacer(1, 1.5 * mm))
    story.append(P("Agbowa, Pacific Court and Gorge View Court are on the earlier list and were not in the November registers. Confirm each is still live before anyone posts to it. Do not merge Gorge View Court with GVC Gbagada. A corner piece is an item, not a project. Co-ownership takes the project the contract names, and account 2025. It does not take a project called Co-ownership.", styles["body"]))

    story.append(P("9.  Customers and vendors", styles["h1"]))
    story.append(P("Do not import the client workbook. It holds phone numbers, email addresses, dates of birth and, in the commission file, agents' bank account numbers. Create the fields the posting needs: Register status, Plot or unit, and Payment plan. Currency NGN. Terms: due on receipt. Create customers by hand, or from a file that contains only the name, the project and the plot. An agent is a vendor. Commission is account 6035, never a deduction from the receipt and never a deduction from revenue.", styles["body"]))

    story.append(P("10.  Bank accounts", styles["h1"]))
    story.append(pdf_table(styles, ["Code", "Account", "Role"], [
        ["1010", "Providus Bank", "Operations"],
        ["1020", "First Bank of Nigeria", "Collections"],
        ["1030", "Sterling Bank", "Disbursements"],
        ["1040", "Guaranty Trust Bank", "The fourth company account"],
    ], [22 * mm, 62 * mm, 94 * mm]))
    story.append(Spacer(1, 2 * mm))
    for n, line in enumerate([
        "Do not create these again from the banking screen. A second account of the same name is how a transfer gets posted twice.",
        "Connect a feed where the bank offers one. Otherwise save a statement-import map and import one week. Read five lines before you import history. A swapped day and month will import, and it will be wrong.",
        "Turn automatic categorisation off. Zoho will offer to treat a subscriber deposit as income or as a customer payment. Both are wrong.",
        "A bank rule may only suggest a match to a retainer payment or a bill payment. It must not create income, an expense, or a customer payment.",
        "A movement between the four banks uses the Transfer action. Do not create a transfer account. Petty cash is 1050 and site imprest is 1055. Do not connect a feed to either.",
    ], 1):
        story.append(P(f"{n}.  {line}", styles["step"]))

    story.append(P("11.  Opening balances", styles["h1"]))
    for n, line in enumerate([
        "Settings, Opening Balances. Date: 31 December 2025. Enter each signed line against the account in this chart. Cash is the four banks, not one bank.",
        "Subscriber money held at that date, not yet allocated and not yet handed over, is 2020. It is not revenue and not a receivable.",
        "Co-ownership principal held at that date is 2025. Do not park it in 2020. The two liabilities are different promises.",
        "Opening Balance Offset 3098 and Opening Balance Adjustments 2090 must both be nil. Do not go live with a balance in either.",
        "Do not enter 2026 transactions as opening balances. If a signed line has no account, add it in the chart file and re-import that row. Do not drop it into General Income.",
    ], 1):
        story.append(P(f"{n}.  {line}", styles["step"]))

    story.append(P("12.  Numbers and templates", styles["h1"]))
    story.append(P("Set the prefixes before the first document is saved: RET-, INV-, CN-, PMT-, BILL-, VPMT-, JNL-, each starting at 0001. Client Relations may raise a retainer and record a payment. Only Finance may raise an invoice, a credit note, a bill or a journal. On the retainer template, show the project, the plot, the amount received, the contract amount and the balance. Hide shipping. State the 10 per cent default charge in the terms, and nothing stronger. Do not print a co-owner's return rate on a land receipt.", styles["body"]))

    story.append(P("13.  Users and roles", styles["h1"]))
    story.append(P("Do not share a login. Head of Accounts can do everything. Finance can post and lock, and cannot change the chart. Client Relations can open a subscriber, raise a retainer, record a payment and set the status, and cannot categorise a bank line or publish a journal. Sales can see their project's customers and the balance, and cannot take a payment. Turn on two-factor authentication before you send the invitations. Remove the consultant's admin access at hand-over, or date the support window.", styles["body"]))

    story.append(P("14.  Preferences that change a posting", styles["h1"]))
    story.append(P("Settings, Preferences. Work down the list. The defaults Zoho ships will post a subscriber deposit as income if you leave them.", styles["body"]))
    for line in [
        "Retainer invoices on. The retainer liability account is 2020. If it points at Sales, every receipt becomes income.",
        "Invoice-on-payment off for land and housing. The invoice is raised on allocation or handover, and the retainer is applied to it.",
        "Automatic late fees off. The 10 per cent default charge is raised by a person, and only when an instalment was missed.",
        "Revenue recognition schedules off. A date passing must not turn a receipt into a sale.",
        "Shipping off. Manual journals on, restricted to Finance. Audit trail on. Multi-currency off unless a foreign bill has arrived, in which case differences go to 9010.",
        "Transaction locking on. Lock to 31 December 2025 once the opening balances tie, and prove that a test user cannot save a 2025 journal.",
    ]:
        story.append(P("•  " + line, styles["bullet"]))

    story.append(P("15.  Reports, the test, and the hand-over", styles["h1"]))
    story.append(P("Save a report of 2020 by project, a report of 2025 beside 2060, a report of open retainers, a report of 1180 and 2045, and the trial balance. Schedule the trial balance and the subscriber-funds report to the Head of Accounts each Monday. Do not schedule a report that contains phone numbers or email addresses.", styles["body"]))
    story.append(P("Before anyone is trained, on a customer named Test Subscriber and on project Head Office: raise a retainer for ₦1,000, record the payment to Providus, void both, and confirm 2020 and the bank are back where they started. Publish a journal for ₦1 and reverse it. If any step posts to Sales, stop and correct the item or the retainer account. Do not run the test on a real subscriber, and do not run it in a locked period.", styles["body"]))
    story.append(P("Train Client Relations and Sales on the retainer, the payment and the status only. Train Finance on the bank match, the allocation, the handover, the co-ownership accrual and the lock. Give each room the posting guide. This guide stays with the Head of Accounts.", styles["body"]))

    story.append(P("16.  Sign-off", styles["h1"]))
    for line in [
        "Premium confirmed. Trial end date written down if the subscription is unpaid.",
        "Nigeria, NGN, January to December, Africa/Lagos. Address read aloud from a printed retainer.",
        "Chart imported with Overwrite. 136 active accounts. Account ID not mapped.",
        "2020, 2025, 2060, 4042, 4120 and 7020 opened and read. No VAT on land, housing or retainers.",
        "Projects created, including Head Office and Ibadan Office. Retainer liability is 2020.",
        "Automatic categorisation off. Opening balances tie. 3098 and 2090 nil. 2025 locked.",
        "Test receipt raised and voided. Two-factor authentication on. Consultant admin access dated.",
    ]:
        story.append(P("•  " + line, styles["bullet"]))
    story.append(Spacer(1, 3 * mm))
    story.append(P("Consultant: ______________________________    Date: ________________", styles["body"]))
    story.append(P("Head of Accounts: ________________________    Date: ________________", styles["body"]))
    story.append(P("Director: ________________________________    Date: ________________", styles["body"]))
    story.append(P("Prepared for Baay Projects Limited. Account codes are the chart imported in section 5. Posting steps are in the posting guide of the same date. This setup is not a tax ruling and not an audit opinion.", styles["small"]))

    doc = SimpleDocTemplate(
        str(PDF),
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=18 * mm,
        bottomMargin=16 * mm,
        title="Baay Projects Limited — Zoho Books Premium setup guide",
        author="Baay Projects Limited",
    )
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)


if __name__ == "__main__":
    build_docx()
    build_pdf()
    print(DOCX)
    print(PDF)
