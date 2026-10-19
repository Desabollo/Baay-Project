"""Engagement proposal: Zoho Books for Baay Projects Limited.

Writes an editable Word copy and an unlocked PDF with a fillable acceptance.
Figures in the narrative are taken from the 2023–2025 master workbook, the
sales and customer register, and the signed financial statements.
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Twips
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
)

ROOT = Path(__file__).resolve().parents[1]
OUT_DOCX = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Engagement_Proposal.docx"
OUT_PDF = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Engagement_Proposal.pdf"

TEAL = "0E4D5C"
TEAL_MID = "1F6F78"
INK = "1C2830"
MUTED = "5E6A72"
GOLD = "B08968"
CREAM = "F6F1E8"
PALE = "F4F7F8"
LINE = "D5DDE1"
WHITE = "FFFFFF"
RED_BROWN = "8C4A3A"

FEE = 1_800_000
VAT_RATE = 0.075
VAT = int(FEE * VAT_RATE)
TOTAL = FEE + VAT

REF = "CEA/BPL/ZB/2026-01"
DATE = "19 October 2026"
VALID = "18 November 2026"


def naira(amount, decimals=0):
    if decimals:
        return f"₦{amount:,.{decimals}f}"
    return f"₦{amount:,.0f}"


def rgb(hex_color):
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def rl(hex_color):
    h = hex_color.lstrip("#")
    return Color(int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255)


# ---------------------------------------------------------------------------
# Shared narrative
# ---------------------------------------------------------------------------

def blocks():
    """Body after the cover and the letter. Rendered by both Word and PDF."""
    return [
        ("h1", "1.  The decision, on one page"),
        ("p", "Baay Projects Limited collects from subscribers long before a plot is allocated or a house is handed over. The cash sits in four banks. The subscriber sits in a client register. The sale, when it is a sale, has to sit in a ledger. Those three things are not the same record today. This engagement puts them in one."),
        ("p", "We will set up Zoho Books, Nigeria edition, in the Company’s name, and teach the people who already sell and bank the money how to use it. We will open the books from the statement of financial position you signed at 31 December 2025, so 2026 starts from a known place rather than from another spreadsheet."),
        ("table",
         ["Item", "Amount"],
         [
             ["Professional fees for this setup", naira(FEE)],
             ["Value added tax at 7.5%", naira(VAT)],
             ["Amount payable to CloudERP", naira(TOTAL)],
         ],
         [0.62, 0.38],
         {2}),
        ("note", "The professional fee is ₦1.8 million, exclusive of VAT, and exclusive of the software subscription. We do not mark up Zoho. We are not asking you to buy an ERP."),
        ("bullets", [
            "Eight weeks from the date we both sign and the first invoice is paid.",
            "Zoho Books Premium, billed annually, is ₦139,200 a year before local taxes, on Zoho’s Nigeria price list at the date of this proposal. That is the Company’s contract with Zoho, not with us.",
            "You will be able to see, in one place, cash by bank, collections by estate, and money received that is not yet a sale.",
            "Housing revenue will not recognise itself. A handover is a decision, recorded by your people, and then the books follow.",
        ]),

        ("h1", "2.  How the business actually runs"),
        ("p", "This section is here so you can see that the proposal is about Baay, not about software in general. It is drawn from the Company’s master workbook for 2023 to 2025, the sales and customer register, and the four bank accounts."),
        ("h2", "What you sell"),
        ("p", "The Company was incorporated on 18 September 2018 and has traded as Baay Projects Limited since 28 June 2021. The registered objects are wide. The business the books have to serve is narrower, and more demanding: land and housing, sold on instalments, plus other trading beside the estates."),
        ("p", "Most subscribers are on a three-month plan. A smaller number pay outright, or over six or twelve months. The client register holds 214 subscriptions, about 190 subscribers, across land and housing. Between 2023 and 2025 the sales ledger recorded 859 receipts."),
        ("table",
         ["Estate", "How it should live in the books"],
         [
             ["Green City, Ketu-Epe and Omu-Epe", "Land, sold in phases at Lantaba, Idobi and Omu. Allocation is the moment of sale."],
             ["Green City Ibadan", "A smaller land scheme at Akinyele-Moniya. Same rule, own project."],
             ["Heritage Garden and Agbowa", "Residential land. Kept separate so a small scheme is not lost in Green City."],
             ["Pacific Court and Pacific Apartment", "Housing. Receipts are large. The sale is the handover, not the transfer."],
             ["Gorge View Court, Gbagada", "Housing. The largest receipt pool in the sales ledger: ₦780 million."],
             ["Baay Foreshore", "Waterfront land and housing. Own project, own collections."],
         ],
         [0.34, 0.66],
         None),
        ("p", "Ancillary charges — corner piece, documentation, development — are real money and are easy to bury inside a plot price. They will be their own items. Other trading will have its own income and cost lines, so a trade does not look like a plot sale."),
        ("h2", "Cash is not the sale"),
        ("p", "This is the distinction the system has to protect. Over the three years the sales ledger collected ₦2.585 billion from customers. The financial statements recognised ₦1.376 billion of revenue. The difference is not a hole. It is money received for performance that had not yet been completed, and it is how a developer should look."),
        ("table",
         ["", "2023", "2024", "2025"],
         [
             ["Collected from customers", "125,620,000", "594,080,600", "1,865,113,400"],
             ["Revenue recognised", "84,117,680", "350,864,060", "940,893,400"],
             ["Contract liabilities at 31 December", "41,602,320", "289,618,860", "1,223,838,860"],
         ],
         [0.40, 0.20, 0.20, 0.20],
         None),
        ("p", "At 31 December 2025 the Company also held ₦1.02 billion of land, housing stock and development cost, ₦223.6 million of co-ownership liabilities, and ₦155.9 million of cash across the banks. Total assets were ₦3.18 billion. Co-ownership is an investment receipt. It is not revenue, and the system will not be allowed to treat it as revenue."),
        ("h2", "Four banks, and more than one list"),
        ("p", "Cash is held with Providus Bank, First Bank of Nigeria, Sterling Bank and GTBank. Money moves between those accounts. A transfer from one to another is not income and it is not an expense. If the books cannot match it, both the profit and the cash will be wrong, in opposite directions."),
        ("p", "The sales list, the commission records and the client register are not the same file, and none of them is the bank. Client Relations knows the subscriber. Finance knows the narration. The bank knows the naira. Zoho Books is the place those three meet, on the day the money arrives, not three weeks later in a workbook."),
        ("note", "Where a bank narration does not yet say what a receipt or payment was for, the books will park it. It will not be forced into profit to make a month look tidy. Naming it is a weekly job. This engagement builds the place it waits, and the habit of clearing it. It does not include a line-by-line clearance of every historical narration."),

        ("h1", "3.  What we will set up"),
        ("h2", "Why Zoho Books, and why Premium"),
        ("p", "You asked for a cloud accounting system. Zoho Books is the right weight. It is cloud, so a receipt raised in the field and a reconciliation done in the office are the same record. The Nigeria edition works in naira and can carry VAT. It is serious enough for estates, inventory and a fixed-asset register, and it is not an ERP implementation in disguise."),
        ("p", "Standard is too small: three users, no projects, no inventory. Professional can track projects and stock, but it has no revenue recognition and no fixed-asset register, and five users is tight once the Managing Director, Finance, Accounts, Client Relations and the company secretary all need a login. Elite adds warehouses and serial numbers. That is not how a plot is sold. Premium is the fit."),
        ("table",
         ["", "Standard", "Professional", "Premium"],
         [
             ["Annual price, before tax", "₦43,200", "₦91,200", "₦139,200"],
             ["Users included", "3", "5", "10"],
             ["Bank feeds, VAT, reconciliation", "Yes", "Yes", "Yes"],
             ["Projects and inventory", "No", "Yes", "Yes"],
             ["Revenue recognition", "No", "No", "Basic"],
             ["Fixed assets, budgets, cash-flow forecast", "No", "No", "Yes"],
             ["Our recommendation", "", "", "This one"],
         ],
         [0.40, 0.20, 0.20, 0.20],
         {6}),
        ("p", "Prices are Zoho’s Nigeria list at 19 October 2026, billed annually, exclusive of local taxes. Monthly billing is higher: Premium would be ₦167,040 a year before tax. An extra user beyond the ten included is ₦800 a month if you stay on annual billing. Any tax on the subscription is between the Company and Zoho. We will help you open the subscription. We will not sit in the middle of it."),
        ("h2", "How a naira will be treated"),
        ("bullets", [
            "Each live estate is a project. A new phase can be added without a new system.",
            "A subscriber is a customer, with the plot or unit, the plan, and the status your people already use: Paying, Bought, Paid, handed over, Refund.",
            "A receipt before allocation, or before a housing handover, credits the subscriber and credits customer deposits. It does not credit sales.",
            "Land becomes revenue when control has transferred — the plot is allocated or the register says Bought — and not above the balance on that subscriber’s contract.",
            "Housing becomes revenue only when handover is recorded. The software will not do this quietly in the background. A status change prompts Finance to post the recognition.",
            "A refund reverses the deposit or the sale it relates to. It is not an expense dressed up as one.",
            "Co-ownership credits a liability. It never credits sales.",
            "A movement between the Company’s own accounts is a transfer. It cancels.",
            "A receipt or payment whose nature is not yet known waits in a holding account, outside profit.",
            "Directors and related parties have their own current account. Their money is not subscriber money.",
            "Land, housing stock and development cost sit in inventory, and move to cost of sales only when the related revenue is recognised.",
            "Equipment is capitalised and depreciated at 20 per cent a year, straight line, a full year in the year of purchase — the policy you already use.",
        ]),
        ("h2", "A week in the new books"),
        ("p", "Monday. Client Relations records a ₦1.5 million instalment on a Green City plot that is still Paying. Zoho credits the subscriber and credits customer deposits. Sales do not move."),
        ("p", "Wednesday. Finance sees the same amount leave Providus and arrive in Sterling. The two lines are matched as a transfer. Profit does not move."),
        ("p", "Friday. A house at Gorge View Court is handed over. Client Relations records the handover. Finance releases the deposit to revenue and releases the related cost. That is the sale."),
        ("p", "Month-end. The Managing Director opens one screen: cash by bank, collections by estate, deposits that are not yet sales, and co-ownership on its own line. The month is then locked. A locked month stays locked."),
        ("h2", "Who sees what"),
        ("table",
         ["Person", "What they can do"],
         [
             ["Managing Director", "See the estate, cash and deposit screens. Approve, rather than re-key."],
             ["Finance and Accounts", "Post, reconcile, recognise revenue, lock a period, and pull the month-end pack."],
             ["Client Relations", "Open a subscriber, record a receipt and a status. Cannot change the chart of accounts or reopen a month."],
             ["Company secretary", "Read. Enough to answer a director, not enough to post."],
             ["CloudERP, during the engagement", "Configure and support. Admin rights return to the Company at the end of hypercare, unless you ask us to stay."],
         ],
         [0.32, 0.68],
         None),
        ("p", "Residential land and housing will be set up so that VAT is not charged on them, which is the Company’s existing treatment. A 7.5 per cent code will be available for supplies that are taxable, including other trading where that is the right answer. We will confirm the codes with you before go-live. We are not giving a tax ruling in this proposal, and we are not filing the returns."),

        ("h1", "4.  Scope"),
        ("h2", "Included in the ₦1.8 million"),
        ("bullets", [
            "A Zoho Books organisation in the Company’s name, Nigeria edition, naira, on the Premium plan.",
            "A chart of accounts that can produce the lines you already report: revenue, deposits, inventory, co-ownership, related parties, cash, tax.",
            "Projects for the estates named in this proposal, and items for plots, housing, ancillary charges, trading and co-ownership.",
            "Customer and vendor opening lists from the client register, cleaned of obvious duplicates.",
            "The four bank accounts, and a test of whether each bank will feed. Where a feed is not available, a weekly statement import instead. We will not pretend a feed exists.",
            "Opening balances that agree to the signed statement of financial position at 31 December 2025, including cash of ₦155,904,157.58.",
            "A worked posting of each transaction type with Finance: a plot instalment, a housing instalment, a handover, a co-ownership receipt, a refund, an interbank transfer, and a payment not yet identified.",
            "The five screens in the month-end pack described above.",
            "User roles, and a period lock.",
            "Three training sessions, for up to eight people: Finance and Accounts, Client Relations, and the Managing Director.",
            "A short procedure note, in Word, that your team keeps. It will say what to post, who posts it, and what never goes to sales.",
            "Thirty days of hypercare after go-live, including sitting with Finance for the first month-end.",
            "Travel within Lagos for the kick-off and the training. We will come to Agege. We will also work remotely, so the engagement is not spent in traffic.",
        ]),
        ("h2", "Not included"),
        ("bullets", [
            "The Zoho subscription, and any tax Zoho charges on it.",
            "Value added tax on our fee. That is added, as shown in section 6.",
            "A reload of every 2023 to 2025 bank line. There are some twenty-three thousand of them. Opening from 31 December 2025 is the clean start. A history load is a separate quote, if you later want it.",
            "Clearing, one by one, the receipts and payments whose nature is still unidentified. We build the holding accounts and the weekly routine. The naming is your team’s, with our help on the method during hypercare.",
            "The statutory audit, tax computations, and filings with the Nigeria Revenue Service or the Corporate Affairs Commission.",
            "Payroll processing. Staff costs can be recorded. A payroll system is a later decision.",
            "A second company, a new bank, or an estate that is not named in this proposal and is launched after go-live. A small addition during hypercare we will absorb. A new entity we will not.",
            "Ongoing bookkeeping after the thirty days. We can propose a monthly retainer then. It is not hidden inside this fee.",
        ]),

        ("h1", "5.  Eight weeks"),
        ("p", "The clock starts when both sides have signed and the first invoice has been paid. If you accept in October, go-live can still fall before the 2026 year-end. If acceptance moves into late November, go-live moves into January. Either way, we open from 31 December 2025, so 2026 is not rebuilt by hand."),
        ("table",
         ["Week", "What happens", "You will have"],
         [
             ["1", "Kick-off at your office. We walk the estates, the banks and the people who touch a receipt.", "A named owner on your side, and a list of the decisions only you can make."],
             ["2", "Chart of accounts, projects, items, tax codes and roles.", "A configuration note for Finance to approve. Five working days to comment."],
             ["3", "Subscriber and vendor load. Bank accounts opened. Feeds tested.", "A customer list you recognise, and an honest answer on which banks will feed."],
             ["4", "Opening balances from the signed 31 December 2025 statement of financial position.", "A trial balance that ties. This is the second invoice point."],
             ["5", "Parallel run on a live week of receipts, using the seven transaction types.", "Finance has posted them, not merely watched us post them."],
             ["6", "The month-end screens, and a dry close.", "A pack the Managing Director can open without a workbook."],
             ["7", "Training. Three sessions. The procedure note is handed over.", "People who can run the next Monday without us in the room."],
             ["8", "Go-live. Hypercare starts. Admin rights move to the Company.", "The third invoice point, and a system that is yours."],
         ],
         [0.12, 0.46, 0.42],
         None),
        ("p", "A sign-off that takes longer than five working days moves the following week by the same number of days. We would rather move the date than pretend a step is done."),

        ("h1", "6.  Fees"),
        ("p", "The professional fee is ₦1,800,000 (one million, eight hundred thousand naira), exclusive of VAT, and exclusive of the Zoho Books subscription. VAT is 7.5 per cent under the Nigeria Tax Act 2025. There is no success fee, and there is no separate charge for Lagos travel."),
        ("table",
         ["Milestone", "Fee", "VAT at 7.5%", "Payable"],
         [
             ["On acceptance, before we start", naira(900_000), naira(67_500), naira(967_500)],
             ["When the opening trial balance ties, and you have approved the configuration", naira(540_000), naira(40_500), naira(580_500)],
             ["On go-live, when training has been delivered", naira(360_000), naira(27_000), naira(387_000)],
             ["Total payable to CloudERP", naira(FEE), naira(VAT), naira(TOTAL)],
         ],
         [0.46, 0.18, 0.18, 0.18],
         {3}),
        ("p", "Our invoice will carry the account details. Please do not pay an account that is not on that invoice."),
        ("h2", "The subscription, which is not our fee"),
        ("table",
         ["Zoho Books, Company to Zoho", "Before local taxes"],
         [
             ["Premium, recommended, billed annually", "₦11,600 a month, ₦139,200 a year"],
             ["Premium, if billed monthly", "₦13,920 a month, ₦167,040 a year"],
             ["Extra user beyond 10, annual billing", "₦800 a user a month"],
             ["Professional, if you later prefer the smaller plan", "₦91,200 a year. We do not recommend it."],
         ],
         [0.62, 0.38],
         None),
        ("p", "Source: Zoho Books Nigeria price list, 19 October 2026, which states that prices are exclusive of local taxes. Zoho can change the list. The figure on Zoho’s invoice is the figure you pay them. If the published price has moved when you subscribe, we will tell you before you commit. We are not a reseller, and we do not add a margin."),
        ("p", "This proposal is open for acceptance until 18 November 2026. After that date the fee is no longer held, and we will confirm it again in writing if you still want to proceed."),

        ("h1", "7.  What we need from you"),
        ("bullets", [
            "One owner. We suggest Finance, with Client Relations beside them. A project with two owners and no owner will slip.",
            "The signed financial statements at 31 December 2025, as the opening source. We will not open from a reconstruction.",
            "The client register, and a person who can say, for a sample of bank narrations, whether a line is a plot, a house, a director, co-ownership, or still unknown.",
            "Authority to open the Zoho subscription in the Company’s name, and to pay Zoho directly.",
            "View access to the four internet-banking profiles for the feed test, or a statement export where a feed is refused. We do not need to be signatories, and we should not be.",
            "The people named for training, in the room or on a call, for the three sessions. Training a deputy who then trains nobody else is not training.",
            "Comments on each sign-off within five working days.",
        ]),
        ("p", "We will not ask you to send passwords by email. Access is granted to named people, and it is withdrawn when the engagement ends, unless a later retainer says otherwise."),

        ("h1", "8.  Terms"),
        ("bullets", [
            "Once signed, this proposal is the engagement letter. If there is a conflict between a later conversation and this document, this document wins until we both change it in writing.",
            "Either side may stop the engagement on fourteen days’ written notice. Fees for milestones already completed remain payable. An advance that has not been earned is returned.",
            "The Company owns the Zoho organisation and everything in it. We configure it. We do not keep it.",
            "We will keep your records confidential, limit access to named staff, and not use the Company’s name as a public case study without your written consent.",
            "We are not your statutory auditor. That appointment is unchanged by this engagement. Nothing in this proposal is an audit opinion, a tax opinion, or a promise that a filing will be accepted.",
            "We do not guarantee a particular profit, a particular tax outcome, or that every Nigerian bank will offer a live feed into Zoho.",
            "Our liability under this engagement is limited to the professional fees you have paid us under it, except where the law does not allow a limit. We do not exclude liability for fraud.",
            "The engagement is governed by the laws of the Federal Republic of Nigeria. The courts of Lagos have jurisdiction.",
        ]),

        ("h1", "9.  Acceptance"),
        ("p", "By signing, the Company accepts this proposal, including the professional fee of ₦1,800,000 exclusive of VAT, the exclusion of the Zoho Books subscription, and the scope in sections 4 and 5."),
        ("p", "One authorised director is enough. A second signature is welcome, and is the way the Board already signs."),
    ]
