"""Revised annual financial statements and the basis-of-revision report."""

from __future__ import annotations

import pickle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
    KeepTogether, HRFlowable, ListFlowable, ListItem,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

NAVY = colors.HexColor("#1F4E79")
GOLD = colors.HexColor("#C4A35A")
PALE = colors.HexColor("#F4F7FB")
LINE = colors.HexColor("#D0D7E2")
RED = colors.HexColor("#8C2F2F")
GREEN = colors.HexColor("#1E6B4F")

OLD = {
    2023: {"revenue": 38_500_000, "cash": 2_842_403, "receivables": 5_015_000,
           "deferred": 2_400_000, "assets": 20_363_943, "pbt": 7_300_000},
    2024: {"revenue": 64_200_000, "cash": 5_142_678, "receivables": 8_685_000,
           "deferred": 3_600_000, "assets": 35_187_218, "pbt": 12_500_000},
    2025: {"revenue": 112_500_000, "cash": 8_924_048, "receivables": 14_360_000,
           "deferred": 5_800_000, "assets": 60_158_588, "pbt": 22_600_000},
}


def N(x, dash_zero=True):
    if x is None:
        return "—"
    x = float(x)
    if dash_zero and abs(x) < 0.005:
        return "—"
    s = f"{abs(x):,.2f}"
    return f"({s})" if x < 0 else s


def Nm(x):
    """Millions, one decimal, for narrative."""
    x = float(x)
    sign = "-" if x < 0 else ""
    return f"{sign}₦{abs(x)/1_000_000:,.1f} million"


class RevCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved = []

    def showPage(self):
        self._saved.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        n = len(self._saved)
        for state in self._saved:
            self.__dict__.update(state)
            self._draw(n)
            super().showPage()
        super().save()

    def _draw(self, n):
        self.saveState()
        if self._pageNumber > 1:
            self.setFillColor(NAVY)
            self.rect(0, A4[1] - 28, A4[0], 28, fill=1, stroke=0)
            self.setFillColor(colors.white)
            self.setFont("DejaVuBold", 7.5)
            self.drawString(36, A4[1] - 18, getattr(self, "header_left", "BAAY PROJECTS LIMITED  ·  RC 1526224"))
            self.setFont("DejaVu", 7.5)
            self.drawRightString(A4[0] - 36, A4[1] - 18, getattr(self, "header_right", "REVISED"))
            self.setFillColor(colors.HexColor("#666666"))
            self.setFont("DejaVu", 7.5)
            self.drawString(36, 18, "Revised financial statements — prepared from bank, sales and client records. Not a signed audit opinion.")
            self.drawRightString(A4[0] - 36, 18, f"Page {self._pageNumber} of {n}")
        self.restoreState()


def styles():
    ss = getSampleStyleSheet()
    ss.add(ParagraphStyle("CoverCo", fontName="DejaVuBold", fontSize=11, textColor=NAVY, alignment=TA_CENTER, leading=14))
    ss.add(ParagraphStyle("CoverTitle", fontName="DejaVuBold", fontSize=16, textColor=NAVY, alignment=TA_CENTER, leading=20, spaceAfter=6))
    ss.add(ParagraphStyle("CoverSub", fontName="DejaVu", fontSize=9.5, textColor=colors.HexColor("#333333"), alignment=TA_CENTER, leading=13))
    ss.add(ParagraphStyle("H1", fontName="DejaVuBold", fontSize=11, textColor=NAVY, spaceBefore=8, spaceAfter=4, leading=14))
    ss.add(ParagraphStyle("H2", fontName="DejaVuBold", fontSize=9.5, textColor=NAVY, spaceBefore=8, spaceAfter=3, leading=12))
    ss.add(ParagraphStyle("Body", fontName="DejaVu", fontSize=8, leading=11, alignment=TA_JUSTIFY, textColor=colors.HexColor("#222222")))
    ss.add(ParagraphStyle("BodyLeft", fontName="DejaVu", fontSize=8, leading=11, alignment=TA_LEFT, textColor=colors.HexColor("#222222")))
    ss.add(ParagraphStyle("Small", fontName="DejaVu", fontSize=7.2, leading=9.4, alignment=TA_JUSTIFY, textColor=colors.HexColor("#333333")))
    ss.add(ParagraphStyle("Th", fontName="DejaVuBold", fontSize=7, leading=9, textColor=colors.white, alignment=TA_CENTER))
    ss.add(ParagraphStyle("Td", fontName="DejaVu", fontSize=7, leading=9, textColor=colors.HexColor("#222222")))
    ss.add(ParagraphStyle("TdR", fontName="DejaVu", fontSize=7, leading=9, alignment=TA_RIGHT))
    ss.add(ParagraphStyle("TdB", fontName="DejaVuBold", fontSize=7, leading=9, textColor=NAVY))
    ss.add(ParagraphStyle("TdBR", fontName="DejaVuBold", fontSize=7, leading=9, alignment=TA_RIGHT, textColor=NAVY))
    ss.add(ParagraphStyle("Warn", fontName="DejaVu", fontSize=8, leading=11, textColor=RED, alignment=TA_JUSTIFY))
    ss.add(ParagraphStyle("Center", fontName="DejaVu", fontSize=8, leading=11, alignment=TA_CENTER))
    ss.add(ParagraphStyle("Sig", fontName="DejaVu", fontSize=8, leading=11, alignment=TA_LEFT))
    return ss


S = styles()


def P(text, style="Body"):
    return Paragraph(str(text), S[style])


def tbl(data, col_widths, bold_rows=None, total_rows=None, header=True, text_cols=None):
    bold_rows = set(bold_rows or [])
    total_rows = set(total_rows or [])
    text_cols = set(text_cols or [])
    wrapped = []
    for r, row in enumerate(data):
        line = []
        for c, val in enumerate(row):
            if r == 0 and header:
                line.append(Paragraph(str(val), S["Th"]))
            elif c in text_cols or c == 0:
                line.append(Paragraph(str(val), S["TdB"] if (r in bold_rows or r in total_rows) else S["Td"]))
            elif r in bold_rows or r in total_rows:
                line.append(Paragraph(str(val), S["TdBR"]))
            else:
                line.append(Paragraph(str(val), S["TdR"]))
        wrapped.append(line)
    t = Table(wrapped, colWidths=col_widths, repeatRows=1 if header else 0)
    cmds = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("GRID", (0, 0), (-1, -1), 0.2, LINE),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
    ]
    for c in text_cols:
        cmds.append(("ALIGN", (c, 1), (c, -1), "LEFT"))
    if header:
        cmds.extend([
            ("BACKGROUND", (0, 0), (-1, 0), NAVY),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ])
    for r in range(1, len(data)):
        if r % 2 == 0:
            cmds.append(("BACKGROUND", (0, r), (-1, r), PALE))
    for r in total_rows:
        cmds.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#E7EEF6")))
        cmds.append(("LINEABOVE", (0, r), (-1, r), 0.6, NAVY))
    t.setStyle(TableStyle(cmds))
    return t


def hr():
    return HRFlowable(width="100%", thickness=1, color=NAVY, spaceAfter=6, spaceBefore=2)


def signature_block():
    rows = [
        [Paragraph("______________________________", S["Sig"]),
         Paragraph("______________________________", S["Sig"])],
        [Paragraph("<b>Adegoke Segun Babatunde</b>", S["Sig"]),
         Paragraph("<b>Adegoke Mary Ayoboade</b>", S["Sig"])],
        [Paragraph("Managing Director", S["Small"]),
         Paragraph("Company Secretary", S["Small"])],
        [Paragraph("FRC number to be inserted by the director", S["Small"]),
         Paragraph("Lagos", S["Small"])],
    ]
    t = Table(rows, colWidths=[261, 262])
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
    ]))
    return t


def face_pnl(p):
    """Presentation P&L that ties to stored profit after tax."""
    other = p["other_income"]
    bank = p["bank_charges"]
    if bank < 0:
        other = other + (-bank)
        bank = 0.0
    revenue = p["revenue"]
    cos = p["cos"]
    gp = revenue - cos
    lines = [
        ("Revenue from contracts with customers", revenue, False),
        ("Cost of sales", -cos, False),
        ("Gross profit", gp, True),
        ("Other income", other, False),
        ("Customer refunds", -p["refunds"], False),
        ("Staff costs", -p["staff"], False),
        ("Director remuneration", -p["director_rem"], False),
        ("Selling commissions", -p["commission"], False),
        ("Rent", -p["rent"], False),
        ("Professional fees", -p["professional"], False),
        ("Administrative expenses", -p["admin"], False),
        ("Bank charges", -bank, False),
        ("Depreciation", -p["depreciation"], False),
        ("Impairment of receivables", -p["ecl_expense"], False),
        ("Operating profit", None, True),
        ("Finance income", p["finance_income"], False),
        ("Finance costs", -p["finance_cost"], False),
        ("Profit before taxation", None, True),
        ("Income tax expense", -p["tax"], False),
        ("Profit for the year", None, True),
    ]
    # fill computed totals
    op = gp + other - p["refunds"] - p["staff"] - p["director_rem"] - p["commission"] - p["rent"] - p["professional"] - p["admin"] - bank - p["depreciation"] - p["ecl_expense"]
    pbt = op + p["finance_income"] - p["finance_cost"]
    pat = pbt - p["tax"]
    out = []
    for name, val, bold in lines:
        if name == "Operating profit":
            val = op
        elif name == "Profit before taxation":
            val = pbt
        elif name == "Profit for the year":
            val = pat
        out.append((name, val, bold))
    return out, pat


def sfp_rows(s):
    assets = [
        ("Property, plant and equipment", s.get("ppe_nbv", 0)),
        ("Inventories — land, development and housing stock", s.get("inventory", 0)),
        ("Trade receivables, net of expected credit loss", s.get("receivables_net", 0)),
        ("Customer receipts recorded but not covered by identified bank property receipts", s.get("untraced_receipts", 0)),
        ("Other source receipts not traced to the four banks", s.get("untraced_other", 0)),
        ("Payments pending allocation", s.get("payments_pending", 0)),
        ("Foreign-currency and other advances", s.get("fx_advances", 0)),
        ("Amounts due from related parties", s.get("related_party_dr", 0)),
        ("Loan repayments in excess of proceeds traced", s.get("borrowings_dr", 0)),
        ("Cash and cash equivalents", s.get("cash", 0)),
    ]
    liab = [
        ("Contract liabilities — deferred income", s.get("contract_liabilities", 0)),
        ("Co-ownership investment liabilities", s.get("co_ownership", 0)),
        ("Receipts pending allocation", s.get("receipts_pending", 0)),
        ("Net transfers with company accounts outside the extraction", s.get("interaccount_cr", 0)),
        ("Current tax liabilities", s.get("tax_cr", 0)),
    ]
    eq = [
        ("Share capital", s.get("share_capital", 0)),
        ("Retained earnings", s.get("retained_earnings", 0)),
    ]
    return assets, liab, eq


def cashflow(model, year):
    b = model["bank"]
    sub = b[b["year"] == year]
    def cr(cat):
        return float(sub.loc[sub["category"] == cat, "cr"].sum())
    def dr(cat):
        return float(sub.loc[sub["category"] == cat, "dr"].sum())
    cust = cr("CUSTOMER_COLLECTION")
    other_in = cr("OTHER_INCOME_TRADING") + cr("OTHER_INCOME_SERVICE") + cr("OTHER_INCOME_REFUND") + cr("FINANCE_INCOME") + cr("BANK_CHARGE_CONTRA")
    pending_in = cr("UNALLOCATED_IN")
    project_out = dr("INVENTORY_LAND") + dr("INVENTORY_HOUSING") + dr("INVENTORY_DEVELOPMENT") + dr("TRADING_PURCHASES") - cr("INVENTORY_CONTRA")
    opex_out = (dr("STAFF_COST") + dr("DIRECTOR_REMUNERATION") + dr("SELLING_COMMISSION") + dr("RENT")
                + dr("PROFESSIONAL") + dr("ADMIN_OPS") + dr("BANK_CHARGE") - cr("BANK_CHARGE")
                + dr("CUSTOMER_REFUND") + dr("FINANCE_COST"))
    pending_out = dr("UNALLOCATED_OUT")
    tax_paid = dr("TAX_PAID")
    ppe = dr("PPE")
    fx = dr("FX_ADVANCE")
    rp = cr("RELATED_PARTY_IN") - dr("RELATED_PARTY_OUT")
    bor = cr("BORROWINGS_IN") - dr("BORROWINGS_OUT")
    div = -dr("DIVIDEND")
    co = cr("CO_OWNERSHIP_IN") - dr("CO_OWNERSHIP_OUT")
    inter = cr("INTERBANK") - dr("INTERBANK")
    net = (cust + other_in + pending_in - project_out - opex_out - pending_out - tax_paid
           - ppe - fx + rp + bor + div + co + inter)
    return {
        "Customer property receipts identified in the banks": cust,
        "Other receipts (trading, refunds, interest)": other_in,
        "Receipts pending allocation": pending_in,
        "Payments for land, housing stock and development": -project_out,
        "Payments for staff, commission, rent, refunds and overheads": -opex_out,
        "Payments pending allocation": -pending_out,
        "Tax paid, as identified in narrations": -tax_paid,
        "Net cash from operating activities": cust + other_in + pending_in - project_out - opex_out - pending_out - tax_paid,
        "Purchase of property, plant and equipment": -ppe,
        "Foreign-currency and other advances": -fx,
        "Net cash used in investing activities": -(ppe + fx),
        "Related-party movements (net)": rp,
        "Borrowings and repayments (net)": bor,
        "Dividends paid": div,
        "Co-ownership receipts (net of returns)": co,
        "Transfers with own accounts outside the extraction (net)": inter,
        "Net cash from financing activities": rp + bor + div + co + inter,
        "Net increase / (decrease) in cash": net,
    }


def build_afs(model, year, path):
    comp = year - 1
    p = model["pnl"][year]
    s = model["sfp"][year]
    if comp == 2022:
        pc = {k: 0.0 for k in p}
        sc = {
            "cash": model["opening_cash_total"],
            "share_capital": model["share_capital"],
            "retained_earnings": model["opening_cash_total"] - model["share_capital"],
            "ppe_nbv": 0, "inventory": 0, "receivables_net": 0, "untraced_receipts": 0,
            "untraced_other": 0, "payments_pending": 0, "fx_advances": 0,
            "related_party_dr": 0, "borrowings_dr": 0, "contract_liabilities": 0,
            "co_ownership": 0, "receipts_pending": 0, "interaccount_cr": 0, "tax_cr": 0,
        }
        sc["total_assets"] = sc["cash"]
        sc["total_equity"] = sc["share_capital"] + sc["retained_earnings"]
        sc["total_liabilities"] = 0
    else:
        pc = model["pnl"][comp]
        sc = model["sfp"][comp]

    face, pat = face_pnl(p)
    if comp == 2022:
        face_c = [(n, 0.0, b) for n, _, b in face]
    else:
        face_c, _ = face_pnl(pc)
    # tie check
    if abs(pat - p["pat"]) > 2:
        raise SystemExit(f"P&L face does not tie for {year}: {pat} vs {p['pat']}")

    W = 523
    story = []
    story.append(Spacer(1, 36))
    story.append(P("REVISED", "CoverCo"))
    story.append(Spacer(1, 4))
    story.append(P("BAAY PROJECTS LIMITED", "CoverTitle"))
    story.append(P("RC 1526224  ·  TIN 21548976-0001", "CoverSub"))
    story.append(Spacer(1, 8))
    story.append(hr())
    story.append(P("REVISED ANNUAL FINANCIAL STATEMENTS", "CoverTitle"))
    story.append(P(f"FOR THE YEAR ENDED 31 DECEMBER {year}", "CoverSub"))
    story.append(P(f"With comparative figures at 31 December {comp}", "CoverSub"))
    story.append(Spacer(1, 8))
    story.append(P(
        "Prepared for audit from the company's bank statements, sales ledger and client register, "
        "so that turnover, receivables, deferred income, cash, assets and liabilities reflect the "
        "transactions in those records. These statements replace the draft pack previously circulated, "
        "which used figures that do not agree to the bank or the sales ledger.",
        "CoverSub"))
    story.append(Spacer(1, 10))
    story.append(P("All amounts are in Nigerian Naira (₦).", "Center"))
    story.append(Spacer(1, 14))
    cover_bits = [
        ["Registered office", "No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State"],
        ["Incorporated", "18 September 2018. Former name Baay Degok Nig Ltd, changed 28 June 2021."],
        ["Directors", "Adegoke Segun Babatunde (Managing Director); Adegoke Mary Ayoboade (Company Secretary); and the other directors named in the CAC status report of 26 August 2026."],
        ["Banks in these statements", "Providus 5400724970 · Sterling 0091166190 · First Bank 2033736938 (Baay Degok Nig Ltd) · GTBank 0893079582"],
        ["Appointed auditor", "Sanni Waheed &amp; Co., Chartered Accountants, Yaba, Lagos. The audit report in this pack is a draft for that firm. It is not a signed opinion."],
        ["Framework", "IFRS and CAMA 2020. Tax for these years follows the law in force before 1 January 2026."],
    ]
    story.append(tbl([["Item", "Detail"]] + cover_bits, [120, 403], header=True, text_cols={1}))
    story.append(Spacer(1, 10))
    story.append(P(f"The revision, at 31 December {year}", "H2"))
    story.append(tbl([
        ["", "Previously circulated", "Revised"],
        ["Customer collections (turnover)", "Not reported", N(model["collections"][year])],
        ["Revenue", N(OLD[year]["revenue"]), N(p["revenue"])],
        ["Deferred income", N(OLD[year]["deferred"]), N(s["contract_liabilities"])],
        ["Trade receivables, net", N(OLD[year]["receivables"]), N(s["receivables_net"])],
        ["Cash at bank", N(OLD[year]["cash"]), N(s["cash"])],
        ["Total assets", N(OLD[year]["assets"]), N(s["total_assets"])],
    ], [220, 151, 152], bold_rows={1}, total_rows={6}))
    story.append(Spacer(1, 4))
    story.append(P(
        "The first column is the pack that was circulated and rejected. It did not agree to the sales ledger or to these bank accounts. "
        "Turnover is every customer receipt in the sales ledger. Revenue is the part of that turnover which IFRS 15 allows to be recognised. "
        "The difference is deferred income, not a missing receipt.",
        "Small"))
    story.append(PageBreak())

    # Directors' report
    story.append(P("REPORT OF THE DIRECTORS", "H1"))
    story.append(hr())
    story.append(P(
        f"The directors submit their report with the revised financial statements of Baay Projects Limited "
        f"for the year ended 31 December {year}. The statements are built from the bank statements, the "
        f"sales ledger and the client register. They are presented so the board and the auditor can see "
        f"the company's actual transactions.",
        "Body"))
    story.append(P("What the company actually does", "H2"))
    story.append(P(
        "The CAC objects include livestock and related farming. The transactions in the review window are "
        "different. The sales ledger records land-plot and housing receipts of ₦2,584,814,000 across 2023 to 2025. "
        "The bank also records commodity and other trading receipts, notably cashew, and co-ownership investment "
        "receipts that are not sales of land or houses. These statements report the activities the records show, "
        "and they disclose the difference from the CAC objects.",
        "Body"))
    story.append(P("Why these statements were revised", "H2"))
    story.append(P(
        "The pack previously circulated reported 2023 revenue of ₦38.5 million, 2024 revenue of ₦64.2 million "
        "and 2025 revenue of ₦112.5 million, and cash of about ₦2.8 million, ₦5.1 million and ₦8.9 million. "
        "None of those figures agrees to the sales ledger or to the four bank accounts. Cash at 1 January 2023 "
        "in that pack was ₦52,500. The bank statements show ₦478,080.25. This revision discards those figures.",
        "Body"))
    story.append(P(f"Results for {year} — turnover is shown separately from revenue", "H2"))
    story.append(P(
        "Under IFRS 15, cash received from a customer is not automatically revenue. Plot sales are recognised "
        "when control of an identified plot has transferred. Housing receipts are deferred until handover is "
        "evidenced. The face of the profit or loss therefore starts with revenue recognised. The company's "
        "own turnover — customer collections in the sales ledger — is the line the board asked to see, and it "
        "is reported here and reconciled in Note 3.",
        "Body"))
    story.append(Spacer(1, 4))
    story.append(tbl([
        ["", f"{year} ₦", f"{comp} ₦"],
        ["Customer collections — turnover per the sales ledger", N(model["collections"].get(year, 0)), N(model["collections"].get(comp, 0) if comp != 2022 else 0)],
        ["Revenue recognised (IFRS 15)", N(p["revenue"]), N(pc["revenue"])],
        ["of which land", N(p["rev_land"]), N(pc.get("rev_land", 0))],
        ["of which housing", N(p["rev_hous"]), N(pc.get("rev_hous", 0))],
        ["Deferred income at 31 December", N(s["contract_liabilities"]), N(sc.get("contract_liabilities", 0))],
        ["Trade receivables, net", N(s["receivables_net"]), N(sc.get("receivables_net", 0))],
        ["Cash at bank", N(s["cash"]), N(sc.get("cash", 0))],
        ["Profit before tax, before allocation of unidentified payments", N(p["pbt"]), N(pc.get("pbt", 0))],
    ], [280, 121, 122], bold_rows={1, 2}, total_rows={8}))
    story.append(Spacer(1, 6))
    story.append(P(
        f"Payments of {Nm(s['payments_pending'])} at 31 December {year} are still pending allocation because the "
        f"bank narration does not identify the asset or the expense. They are carried as an asset, not charged "
        f"to profit. If they were operating expenses, the profit above would not stand. Receipts of "
        f"{Nm(s['receipts_pending'])} are likewise pending allocation and are carried as a liability, not as "
        f"revenue. Both schedules are in the revised master workbook. Profit in these statements is therefore "
        f"a profit before that allocation, not a finished measure of performance.",
        "Warn"))
    story.append(P("Dividend", "H2"))
    story.append(P(
        f"Bank narrations describe dividend payments of {N(p['dividends'])} in {year}. Those amounts have been "
        f"treated as distributions already made, not as a proposal in this report. The directors do not recommend "
        f"a further dividend. The payments are disclosed as related-party distributions because they were made "
        f"to accounts connected with Adegoke Segun Babatunde.",
        "Body"))
    story.append(P("Directors and shares", "H2"))
    story.append(P(
        "The CAC status report of 26 August 2026 records issued share capital of ₦100,000,000 in shares of ₦1, "
        "held as to 80% by Adegoke Segun Babatunde, 10% by Adegoke Mary Ayoboade and 10% by Adegoke Denisa. "
        "The cash records for 2023 to 2025 do not contain a ₦99,000,000 capital injection, and the opening cash "
        "at 1 January 2023 was ₦478,080. These statements therefore recognise paid-up capital of ₦1,000,000, "
        "being the amount previously reported as fully paid, and disclose the CAC figure without recognising "
        "the unpaid or untraced balance as an asset or as additional equity. The board should confirm the "
        "allotment dates and the consideration with the auditor.",
        "Body"))
    story.append(P("Auditors", "H2"))
    story.append(P(
        "Sanni Waheed &amp; Co. are the appointed auditors. The opinion page in this document is a draft for "
        "that firm to complete. It is not a signed audit opinion, and it is drafted as a qualified opinion "
        "because opening balances other than cash, the unallocated payments, and the untraced sales receipts "
        "cannot be cleared from the files alone.",
        "Body"))
    story.append(Spacer(1, 8))
    story.append(P("By order of the board", "Sig"))
    story.append(Spacer(1, 14))
    story.append(signature_block())
    story.append(PageBreak())

    # Responsibilities
    story.append(P("STATEMENT OF DIRECTORS' RESPONSIBILITIES", "H1"))
    story.append(hr())
    story.append(P(
        "The Companies and Allied Matters Act 2020 requires the directors to prepare financial statements "
        "that give a true and fair view. In preparing these revised statements the directors are asked to "
        "adopt the following, which is how the statements have been built from the records:",
        "Body"))
    for item in [
        "The sales ledger of 859 receipt lines, totalling ₦2,584,814,000, is the customer-turnover population. It is not ignored, and it is not replaced by a rounded estimate.",
        "Cash is the balance on the four extracted bank accounts. It is not the cash figure from the previously circulated pack.",
        "Revenue follows IFRS 15. Collections on plots whose control has transferred are revenue. Collections on housing before evidenced handover, and on plots still marked paying, are deferred income.",
        "A receivable is recognised only for the unpaid balance of a contract whose revenue has been recognised. Unpaid balances on contracts not yet performed are disclosed as a pipeline. They are not receivables.",
        "Where a bank narration does not support a classification, the amount is left in receipts pending allocation or payments pending allocation. It is not forced into revenue or into profit.",
        "The statements are prepared on a going-concern basis. Cash at 31 December 2025 was ₦155.9 million and the company was still receiving customer collections. Going concern still depends on collection of the deferred pipeline and on allocation of the pending payments.",
    ]:
        story.append(P("•  " + item, "Small"))
    story.append(Spacer(1, 8))
    story.append(P(
        "The directors are responsible for the records and for safeguarding the assets. These revised statements "
        "were prepared from the files supplied. They are ready for the board to adopt and for the auditor to test.",
        "Body"))
    story.append(Spacer(1, 8))
    story.append(signature_block())
    story.append(Spacer(1, 8))
    story.append(P("Contents of this pack", "H2"))
    story.append(tbl([
        ["Section", "What it is for"],
        ["Directors' report", "The results, the dividend already traced, and the reason the previous figures were withdrawn."],
        ["Draft auditor's report", "Wording for Sanni Waheed &amp; Co. It is not a signed opinion."],
        ["The four statements", "Position, performance, equity and cash. Turnover is reconciled to revenue on the profit or loss."],
        ["Notes 1 to 16", "The policies, the receivables, the deferred income, the pending allocation, and the comparison with the rejected pack."],
        ["Basis of revision", "The companion report for all three years. Read it with this pack."],
        ["Revised workbooks", "The bank extraction, the sales file, the reconciliation and the master model. The statements are drawn from those files."],
    ], [130, 393], text_cols={0, 1}))
    story.append(PageBreak())

    # Draft auditor report
    story.append(P("DRAFT INDEPENDENT AUDITOR'S REPORT", "H1"))
    story.append(P("TO THE MEMBERS OF BAAY PROJECTS LIMITED", "H2"))
    story.append(hr())
    story.append(P(
        "DRAFT FOR COMPLETION BY SANNI WAHEED &amp; CO. THIS PAGE IS NOT A SIGNED AUDIT OPINION AND MUST NOT BE REPRESENTED AS ONE.",
        "Warn"))
    story.append(P("Report on the audit of the revised financial statements", "H2"))
    story.append(P("Qualified opinion — draft", "H2"))
    story.append(P(
        f"We have audited the revised financial statements of Baay Projects Limited, which comprise the statement "
        f"of financial position as at 31 December {year}, and the statement of profit or loss and other comprehensive "
        f"income, statement of changes in equity and statement of cash flows for the year then ended, and notes to "
        f"the financial statements, including a summary of significant accounting policies.",
        "Body"))
    story.append(P(
        "In our opinion, except for the possible effects of the matters described in the Basis for qualified opinion "
        "section of our report, the accompanying revised financial statements give a true and fair view of the "
        "financial position of the Company as at 31 December "
        f"{year}, and of its financial performance and its cash flows for the year then ended, in accordance with "
        "IFRS and the requirements of CAMA 2020, and they reflect the customer collections, bank balances and "
        "contract balances in the records examined.",
        "Body"))
    story.append(P("Basis for qualified opinion — draft", "H2"))
    story.append(P(
        "Opening balances at 1 January 2023, other than cash of ₦478,080.25 on the four bank accounts, could not "
        "be reconstructed. Any land held at that date is not in these statements. Cost of sales, and therefore "
        "profit, may be misstated to that extent.",
        "Small"))
    story.append(P(
        f"Payments of {Nm(model['sfp'][2025]['payments_pending'])} at 31 December 2025, and the corresponding "
        f"balances at the earlier year ends, could not be allocated from the narrations. They are presented as "
        f"payments pending allocation rather than as expenses or inventory. Profit is stated before that allocation.",
        "Small"))
    story.append(P(
        f"Sales-ledger receipts of {Nm(model['sfp'][2025]['untraced_receipts'])} at 31 December 2025 exceed the "
        f"property receipts identified in the four bank accounts on a cumulative basis. They are presented as "
        f"customer receipts not covered by identified bank property receipts. Some of them may sit inside the "
        f"receipts pending allocation. They have not been confirmed to another bank account.",
        "Small"))
    story.append(P(
        "We conducted our audit in accordance with International Standards on Auditing. We are independent of the "
        "Company in accordance with the IESBA Code. The evidence we have obtained is sufficient to support the "
        "opinion above, subject to the qualifications stated. This paragraph is draft wording for the engagement "
        "partner. It is not a representation that the audit procedures have been completed.",
        "Body"))
    story.append(P("Emphasis of matter — revision of previously circulated figures", "H2"))
    story.append(P(
        "We draw attention to Note 1. These statements replace a draft pack whose revenue and cash did not agree "
        "to the sales ledger or the bank statements. Our draft opinion is not modified further in respect of that "
        "replacement, but users should not rely on the previously circulated figures.",
        "Body"))
    story.append(P("Key audit matters — draft", "H2"))
    story.append(P(
        "Revenue recognition and deferred income. The sales ledger is ₦2.58 billion. Most housing receipts remain "
        "deferred because the client register does not evidence handover. The risk is that deferred income is "
        "released too early, or that plot receipts are recognised without control having transferred.",
        "Small"))
    story.append(P(
        "Completeness of cash and classification of narrations. Four accounts were extracted and the running "
        "balances agree. A large residual of narrations could not be classified. The risk is misclassification "
        "between revenue, deferred income, related-party balances and expenses.",
        "Small"))
    story.append(Spacer(1, 10))
    story.append(P("Sanni Waheed &amp; Co.", "Sig"))
    story.append(P("Chartered Accountants", "Sig"))
    story.append(P("1st Floor, No. 29 Olonode Street, Alagomeji, Yaba, Lagos", "Sig"))
    story.append(Spacer(1, 8))
    story.append(P("Engagement partner: name and FRC/2016/ICAN/2016/00000013886 to be signed by the partner.", "Sig"))
    story.append(P("Date: to be inserted on completion of the audit. This draft is not dated as a completed opinion.", "Sig"))
    story.append(PageBreak())

    # SFP
    story.append(P(f"STATEMENT OF FINANCIAL POSITION AS AT 31 DECEMBER {year}", "H1"))
    story.append(hr())
    note_of = {
        "Property, plant and equipment": "8",
        "Inventories": "8",
        "Trade receivables": "5",
        "Customer receipts recorded": "7",
        "Other source receipts": "11",
        "Payments pending": "9",
        "Foreign-currency": "8",
        "Amounts due from related": "10",
        "Loan repayments": "10",
        "Cash and cash": "7",
        "Contract liabilities": "4",
        "Co-ownership": "11",
        "Receipts pending": "9",
        "Net transfers": "7",
        "Current tax": "12",
        "Share capital": "13",
        "Retained earnings": "13",
    }
    def note_for(name):
        for key, num in note_of.items():
            if name.startswith(key):
                return num
        return ""
    assets, liab, eq = sfp_rows(s)
    assets_c, liab_c, eq_c = sfp_rows(sc)
    rows = [["", "Note", f"31 Dec {year}", f"31 Dec {comp}"]]
    rows.append(["ASSETS", "", "", ""])
    bold = {1}
    totals = set()
    for i, ((name, val), (_, valc)) in enumerate(zip(assets, assets_c)):
        if abs(val) < 0.5 and abs(valc) < 0.5:
            continue
        rows.append([name, note_for(name), N(val), N(valc)])
    rows.append(["Total assets", "", N(s["total_assets"]), N(sc.get("total_assets", sc.get("cash", 0)))])
    totals.add(len(rows) - 1)
    rows.append(["LIABILITIES", "", "", ""])
    bold.add(len(rows) - 1)
    for (name, val), (_, valc) in zip(liab, liab_c):
        if abs(val) < 0.5 and abs(valc) < 0.5:
            continue
        rows.append([name, note_for(name), N(val), N(valc)])
    rows.append(["Total liabilities", "", N(s["total_liabilities"]), N(sc.get("total_liabilities", 0))])
    totals.add(len(rows) - 1)
    rows.append(["EQUITY", "", "", ""])
    bold.add(len(rows) - 1)
    for (name, val), (_, valc) in zip(eq, eq_c):
        rows.append([name, note_for(name), N(val), N(valc)])
    rows.append(["Total equity", "", N(s["total_equity"]), N(sc.get("total_equity", sc.get("share_capital", 0) + sc.get("retained_earnings", 0)))])
    totals.add(len(rows) - 1)
    rows.append(["Total liabilities and equity", "", N(s["total_liabilities"] + s["total_equity"]), N(sc.get("total_liabilities", 0) + sc.get("total_equity", sc.get("cash", 0)))])
    totals.add(len(rows) - 1)
    story.append(tbl(rows, [300, 36, 93, 94], bold_rows=bold, total_rows=totals))
    story.append(Spacer(1, 4))
    story.append(P(
        f"The statement balances. Assets of {N(s['total_assets'])} equal liabilities and equity. "
        "Cash equals the four bank statements to the kobo. Amounts on this page are in naira and kobo.",
        "Small"))
    story.append(Spacer(1, 6))
    story.append(P("Cash, by bank account", "H2"))
    story.append(_cash_table(model, year))
    story.append(PageBreak())

    # P&L
    story.append(P(f"STATEMENT OF PROFIT OR LOSS AND OTHER COMPREHENSIVE INCOME FOR THE YEAR ENDED 31 DECEMBER {year}", "H1"))
    story.append(hr())
    story.append(P(
        "Customer collections (the company's turnover) are reconciled to revenue immediately below the statement. "
        "They are not the same number, and both are shown so that neither is hidden.",
        "Small"))
    prow = [["", f"Year ended 31 Dec {year}", f"Year ended 31 Dec {comp}"]]
    bold = set()
    for i, ((name, val, is_bold), (_, valc, _)) in enumerate(zip(face, face_c), start=1):
        prow.append([name, N(val), N(valc)])
        if is_bold:
            bold.add(i)
    story.append(tbl(prow, [300, 111, 112], bold_rows=bold, total_rows=bold))
    story.append(Spacer(1, 4))
    story.append(P("Other comprehensive income is nil. Total comprehensive income equals profit for the year.", "Small"))
    story.append(Spacer(1, 6))
    story.append(P("Turnover reconciliation — sales ledger to revenue recognised", "H2"))
    col = model["collections"].get(year, 0)
    col_c = model["collections"].get(comp, 0) if comp != 2022 else 0
    d_cl = s["contract_liabilities"] - sc.get("contract_liabilities", 0)
    d_rec = s["receivables_gross"] - sc.get("receivables_gross", 0)
    if comp == 2022:
        d_cl_c, d_rec_c = 0.0, 0.0
    else:
        sc2 = model["sfp"][comp - 1] if comp > 2023 else {
            "contract_liabilities": 0, "receivables_gross": 0,
        }
        if comp == 2023:
            sc2 = {"contract_liabilities": 0, "receivables_gross": 0}
        d_cl_c = sc.get("contract_liabilities", 0) - sc2.get("contract_liabilities", 0)
        d_rec_c = sc.get("receivables_gross", 0) - sc2.get("receivables_gross", 0)
    story.append(tbl([
        ["", f"{year}", f"{comp}"],
        ["Customer collections per the sales ledger (turnover)", N(col), N(col_c)],
        ["Increase in deferred income on those collections", N(-d_cl), N(-d_cl_c) if comp != 2022 else "—"],
        ["Increase in unpaid balances recognised as receivables", N(d_rec), N(d_rec_c) if comp != 2022 else "—"],
        ["Revenue recognised", N(p["revenue"]), N(pc.get("revenue", 0))],
    ], [300, 111, 112], total_rows={4}))
    story.append(P(
        "The bridge is the IFRS 15 identity: revenue equals collections, plus the movement in receivables, "
        "minus the movement in deferred income, on the sales-ledger population. It ties to the naira in the workings. "
        "Bank property receipts that could not be tied to a sales line are not added to revenue. Where identified "
        "bank property receipts fall short of the sales ledger on a cumulative basis, the shortfall is the untraced "
        "receipts asset, not extra revenue.",
        "Small"))
    story.append(PageBreak())

    # SOCE
    story.append(P(f"STATEMENT OF CHANGES IN EQUITY FOR THE YEAR ENDED 31 DECEMBER {year}", "H1"))
    story.append(hr())
    if comp == 2022:
        open_sc, open_re = model["share_capital"], model["opening_cash_total"] - model["share_capital"]
    else:
        open_sc, open_re = model["sfp"][comp]["share_capital"], model["sfp"][comp]["retained_earnings"]
    story.append(tbl([
        ["", "Share capital", "Retained earnings", "Total"],
        [f"At 1 January {year}", N(open_sc), N(open_re), N(open_sc + open_re)],
        ["Profit for the year", "—", N(p["pat"]), N(p["pat"])],
        ["Dividends traced in the bank", "—", N(-p["dividends"]), N(-p["dividends"])],
        [f"At 31 December {year}", N(s["share_capital"]), N(s["retained_earnings"]), N(s["total_equity"])],
    ], [180, 110, 116, 117], total_rows={4}))
    story.append(Spacer(1, 6))
    story.append(P(
        "Opening equity at 1 January 2023 equals opening cash of ₦478,080.25. Share capital recognised is "
        "₦1,000,000. The balancing figure is an opening retained-earnings deficit of ₦521,920. No other opening "
        "asset or liability was supported by the files. See Note 13.",
        "Small"))
    story.append(Spacer(1, 12))

    # Cash flow — kept with the equity statement so that page is not left half empty
    story.append(P(f"STATEMENT OF CASH FLOWS FOR THE YEAR ENDED 31 DECEMBER {year}", "H1"))
    story.append(P("Direct method, because the cash book is the bank extraction.", "Small"))
    story.append(hr())
    cf = cashflow(model, year)
    opening_cash_chk = model["opening_cash_total"] if year == 2023 else model["sfp"][year - 1]["cash"]
    if abs((opening_cash_chk + cf["Net increase / (decrease) in cash"]) - s["cash"]) > 1:
        raise SystemExit(
            f"Cash flow does not tie for {year}: "
            f"{opening_cash_chk + cf['Net increase / (decrease) in cash']} vs {s['cash']}"
        )
    if comp != 2022:
        cfc = cashflow(model, comp)
    else:
        cfc = {k: 0 for k in cf}
    opening_cash = model["opening_cash_total"] if year == 2023 else model["sfp"][year - 1]["cash"]
    cf_rows = [["", f"{year}", f"{comp if comp != 2022 else '2022'}"]]
    bold = set()
    for i, (name, val) in enumerate(cf.items(), start=1):
        cf_rows.append([name, N(val), N(cfc.get(name, 0))])
        if name.startswith("Net "):
            bold.add(i)
    cf_rows.append([f"Cash at 1 January {year}", N(opening_cash), "—"])
    cf_rows.append([f"Cash at 31 December {year}", N(opening_cash + cf["Net increase / (decrease) in cash"]), "—"])
    bold.add(len(cf_rows) - 1)
    story.append(tbl(cf_rows, [300, 111, 112], bold_rows=bold))
    story.append(Spacer(1, 4))
    story.append(P(
        f"Cash at 31 December {year} is the sum of the stated bank balances: "
        + ", ".join(f"{bk} ₦{bal:,.2f}" for bk, bal in model["stated_cash"][year]["by_bank"].items())
        + f". Total ₦{model['stated_cash'][year]['total']:,.2f}. The movement agrees to the extraction. "
        "Inter-account transfers between the four banks are eliminated in the classification; the financing "
        "line is the net amount that does not mirror inside the four accounts.",
        "Small"))
    story.append(PageBreak())

    # Notes
    story.append(P("NOTES TO THE REVISED FINANCIAL STATEMENTS", "H1"))
    story.append(hr())
    notes = notes_for(model, year)
    for title, paras in notes:
        story.append(P(title, "H2"))
        for para in paras:
            if isinstance(para, str):
                story.append(P(para, "Small"))
            else:
                story.append(Spacer(1, 3))
                story.append(para)
                story.append(Spacer(1, 3))
    story.append(Spacer(1, 8))
    story.append(P(
        "These revised statements were prepared on 8 October 2026 from the company's records. "
        f"They are for the board of Baay Projects Limited and for Sanni Waheed &amp; Co.",
        "Small"))

    doc = SimpleDocTemplate(
        path, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=40, bottomMargin=36,
        title=f"Baay Projects Limited — Revised financial statements {year}",
        author="Revised from company records",
    )
    doc.build(story, canvasmaker=RevCanvas)
    print("wrote", path)


def notes_for(model, year):
    p = model["pnl"][year]
    s = model["sfp"][year]
    t = model["tax"][year]
    contracts = model["contracts"]
    cl_h = float(contracts.loc[contracts["product"] == "Housing", f"cl_{year}"].sum())
    cl_l = float(contracts.loc[contracts["product"] == "Land", f"cl_{year}"].sum())
    rec_h = float(contracts.loc[contracts["product"] == "Housing", f"rec_{year}"].sum())
    rec_l = float(contracts.loc[contracts["product"] == "Land", f"rec_{year}"].sum())
    pipe = contracts[contracts[f"cl_{year}"] > 0]
    pipe_contract = float(pipe["contract"].sum())
    pipe_n = int(len(pipe))
    old = OLD[year]
    return [
        ("1. Reporting entity, and why this pack was revised", [
            "Baay Projects Limited (RC 1526224) was incorporated on 18 September 2018. Its former name, Baay Degok Nig Ltd, was changed on 28 June 2021. The First Bank account 2033736938 is still in the former name and is included. Baay Gokes, where it appears, is treated as an operating name of the same company, consistent with the single RC number. No second entity is consolidated.",
            "The registered office on the CAC status report of 26 August 2026 is No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State. An operational address at Admiralty Way, Lekki, is recorded in the entity file and is not the registered office.",
            "The draft pack previously circulated reported revenue of ₦38.5 million, ₦64.2 million and ₦112.5 million and cash balances that do not exist on these bank statements. It also cited bank account numbers that are not the accounts in the extraction. Those figures are withdrawn. The comparison is in Note 15.",
            "The functional currency is the naira. The statements are prepared under IFRS and CAMA 2020, on the historical-cost basis. They are not a cash-basis set of accounts: collections are split between revenue, deferred income and receivables. They are also not a set of estimated round numbers.",
        ]),
        ("2. Accounting policies that decide the numbers", [
            "Revenue — land. A plot subscription marked bought, closed or completed is recognised in the year of subscription, or in the first year of collection if the subscription year is outside 2023–2025. The starting point is the contract value. The receivable is the unpaid balance, but it is not recognised above the balance on the client register. Where the register balance is lower, revenue is reduced by the same amount, so the statements do not show a debt the register does not show. Across 2023 to 2025 that reduction is ₦515,131,000. If the register is incomplete, both revenue and receivables are understated by that amount. It is a point for the auditor, not a hidden asset. Receipts above contract value remain deferred. Subscriptions still marked paying, promise, refund or migrated are not recognised.",
            "Revenue — housing. Off-plan housing is recognised at a point in time, on handover. Handover is taken only where the register status is Bought or Completed and at least half of the contract has been paid, or Paid with register paid-to-date of at least 80% of contract value. Pacific Court entries marked paid with a nil paid-to-date and an unsupported contract value are not recognised. All other housing receipts are deferred income. That is why deferred income is large in 2025: the housing book was still largely unhanded-over on the register.",
            "Uncontracted plot receipts, including estate-not-stated lines in the turnover register, are recognised in the year of receipt. Uncontracted housing receipts are deferred. This is a judgment, disclosed because the turnover register treats those plot lines as sales and the register has no open subscription saying otherwise.",
            "Receivables and expected credit loss. A receivable exists only after revenue is recognised and the customer has not yet paid the recognised amount. A general loss allowance of 5% is held. There is no loss history in the files, so the rate is an estimate, not a matrix fitted to data. Unpaid balances on contracts not yet recognised are not receivables. They are disclosed as a pipeline in Note 6.",
            "Inventories. Payments the narration identifies as land, housing stock (including QC Homes units at Gorge View Court) or development materials are capitalised. Cost of sales is that cost, released in proportion to revenue recognised against revenue plus deferred income of the same project. Cost with no customer activity stays in inventory. Pre-2023 land cost is not in the files, so margin on land sold in the period may be overstated.",
            "Cash. Cash is the stated balance on Providus 5400724970, Sterling 0091166190, First Bank 2033736938 and GTBank 0893079582. GTBank has no 2023 transactions; its 1 January 2024 balance was nil. The extraction's running balances agree. No other account has been added.",
            "Pending allocation. A narration that does not support a classification is not forced. Credits remain receipts pending allocation (a liability). Debits remain payments pending allocation (an asset). Profit is not credited or charged with those amounts.",
            "Tax. Years ended on or before 31 December 2025 are taxed under the law in force before the Nigeria Tax Act 2025, which commenced on 1 January 2026. Company size follows gross turnover: 20% companies income tax between ₦25 million and ₦100 million, 30% above ₦100 million. Tertiary education tax is 3% of assessable profit. Police trust fund levy is 0.005% of profit before tax. NASENI levy of 0.25% of profit before tax is accrued for a large company and is subject to confirmation of sector scope. Capital allowances were not claimed: a qualifying-expenditure register was not in the files. Depreciation and the credit-loss allowance are added back. Minimum tax of 0.5% of turnover was computed and was not the binding charge.",
            "Residential plot and housing sales are treated as outside the scope of VAT, which is the usual treatment of bare land and residential property. No output VAT is recognised. If the tax authority takes a different view of the housing sales, a contingent exposure exists. It is not accrued.",
        ]),
        ("3. Turnover, revenue and cost of sales", [
            f"Sales-ledger turnover for {year} is ₦{model['collections'][year]:,.0f}. Across 2023 to 2025 it is ₦2,584,814,000, from 859 lines: land ₦680,144,000 and housing ₦1,904,670,000. Construction receipts in the source files are nil. The first construction engagement in those files is dated February 2026 and is outside this period.",
            f"Revenue recognised in {year} is ₦{p['revenue']:,.0f}, of which land ₦{p['rev_land']:,.0f} and housing ₦{p['rev_hous']:,.0f}. Cost of sales is ₦{p['cos']:,.0f}, of which commodity trading cost is ₦{p.get('cos_trading', 0):,.0f}.",
            "The difference between turnover and revenue is deferred income, not an omission. Housing collections of about ₦1.90 billion across the three years produce housing revenue only where handover is evidenced. The rest is in contract liabilities.",
            "Other income includes commodity trading receipts (cashew and similar narrations), service income from the exclusions register, and refunds received. Trading is reported separately from property revenue because it is not plot or housing turnover.",
            f"Of 859 sales lines, {int(model['sales']['bank_txn'].notna().sum())} lines totalling ₦{model['traced_sales_amt']:,.0f} were agreed to a specific bank credit of the same amount. The financial statements do not use that line match as the asset. They use the cumulative gap between the sales ledger and all bank credits identified as property receipts, so that a receipt which is in the bank but not matched line-by-line is not also booked as a second asset. The cumulative gap at 31 December {year} is ₦{s['untraced_receipts']:,.0f}.",
        ]),
        ("4. Deferred income", [
            f"Contract liabilities at 31 December {year} are ₦{s['contract_liabilities']:,.0f}. Of the contract-level balance, housing is ₦{cl_h:,.0f} and land is ₦{cl_l:,.0f}.",
            "Deferred income is money the company has recorded as received from customers for performance that the register does not yet show as complete. It is a liability. It will become revenue when the plot is allocated and bought, or when the housing unit is handed over, and it will fall if the customer is refunded.",
            f"Co-ownership receipts are not in this line. They are a separate liability of ₦{s['co_ownership']:,.0f}. The source files describe them as a fixed-return investment product, not a sale of land or housing. Returns paid out have been deducted.",
        ]),
        ("5. Trade receivables", [
            f"Gross trade receivables at 31 December {year} are ₦{s['receivables_gross']:,.0f}. The 5% loss allowance is ₦{s['ecl']:,.0f}. Net receivables are ₦{s['receivables_net']:,.0f}. Land accounts for ₦{rec_l:,.0f} of the gross balance and housing for ₦{rec_h:,.0f}.",
            "The balance is small relative to turnover because revenue is recognised only when the contract is substantially collected or the register shows it bought, and the unpaid tail on those contracts is short. The large unpaid housing book is not a receivable. It is deferred, and the amount still to be collected on it is the pipeline in Note 6.",
        ]),
        ("6. Contract pipeline — unpaid balances that are not receivables", [
            f"At 31 December {year}, {pipe_n} customer lines still had deferred income. Their registered contract values total ₦{pipe_contract:,.0f}. Collections on those lines sit in deferred income. The unpaid remainder, where a contract value is recorded, is an amount the company expects to collect if the customer completes the plan. IFRS 15 does not allow that remainder to be recognised as a receivable before the company has an unconditional right to it, which for these contracts is when control transfers.",
            "The client register's own paid-to-date and balance columns were not recomputed. They are the figures typed by client relations. Where they disagree with the sales ledger, the sales ledger is the collection record used in these statements, and the disagreement is a review point, not a silent adjustment.",
        ]),
        ("7. Cash and the four bank accounts", [
            "Account, name on the statement, and balance:",
            _cash_table(model, year),
            "The previously circulated pack cited Providus 5400281942, First Bank 2034891102, Sterling 0078451290 and GTBank 0421897631. Those numbers are not the accounts in the extraction. The accounts above are the ones whose statements were extracted and reconciled, with a nil running-balance difference on the extraction.",
            f"Identified property receipts in these accounts are less than the sales ledger by ₦{s['untraced_receipts']:,.0f} on a cumulative basis at this year end. Receipts pending allocation of ₦{s['receipts_pending']:,.0f} may include further customer money. They have not been moved into turnover without a name, a keyword or a line match.",
        ]),
        ("8. Inventories and property, plant and equipment", [
            f"Inventory at 31 December {year} is ₦{s['inventory']:,.0f}. It is bank payments identified as land, as housing units bought for resale (QC Homes / Gorge View Court and similar), and as development materials, less the portion released to cost of sales. Release is proportional to recognised revenue over recognised revenue plus deferred income of the same project. It is an estimate. It is not a margin that was plugged.",
            f"Property, plant and equipment has a carrying amount of ₦{s['ppe_nbv']:,.0f}. Additions are narrations that identify vehicles, equipment or similar items. Depreciation is 20% straight line, a full year in the year of purchase. The charge in {year} is ₦{p['depreciation']:,.0f}. A physical register was not in the files.",
        ]),
        ("9. Amounts pending allocation — read this with the profit figure", [
            f"Receipts pending allocation: ₦{s['receipts_pending']:,.0f}. Payments pending allocation: ₦{s['payments_pending']:,.0f}. Every item of ₦1 million or more is listed in the revised master workbook and in the revised bank extraction, with the narration.",
            "These are not plugs into revenue and they are not plugs into profit. They are the transactions the narration does not classify. Management and the auditor should clear them to revenue, deferred income, inventory, expenses, related parties or loans. Until that is done, profit is incomplete.",
            f"A sensitivity, not a second set of accounts: charging the entire payments-pending balance to profit would reduce profit before tax for {year} by ₦{s['payments_pending']:,.0f} and would eliminate the reported profit. That sensitivity is why the draft audit opinion is qualified.",
        ]),
        ("10. Related parties", [
            f"Amounts due from related parties at 31 December {year} are ₦{s['related_party_dr']:,.0f}. This is the net of amounts received from and paid to the directors, principally Adegoke Segun Babatunde, after salary and commission have been charged to profit and after amounts labelled dividend have been treated as distributions (₦{p['dividends']:,.0f} in {year}).",
            f"Director remuneration charged in profit is ₦{p['director_rem']:,.0f}. Selling commissions include commissions paid to directors and to realtors; the commission line in total is ₦{p['commission']:,.0f}.",
            "A debit related-party balance means the directors have drawn more than they have contributed in the review window, on the narrations identified. It is recoverable only if the board confirms it is an advance and not a distribution. The auditor should confirm it.",
        ]),
        ("11. Co-ownership", [
            f"Co-ownership liabilities are ₦{s['co_ownership']:,.0f}. The exclusions register removes ₦235.7 million of co-ownership receipts from sales so that they are not turnover. Those receipts, less returns paid, are this liability. Credit memos of ₦613,250 are non-cash and are not posted. Service income in the exclusions register is in other income where it was traced, and in untraced other receipts where it was not.",
        ]),
        ("12. Taxation", [
            f"The company is classified as {t['size']} on turnover of ₦{t['turnover']:,.0f}. Companies income tax at {t['rate']:.0%} is ₦{t['cit_payable']:,.0f}"
            + (" (minimum tax applied). " if t["min_applied"] else ". ")
            + f"Tertiary education tax is ₦{t['tet']:,.0f}. Police trust fund levy is ₦{t['ptf']:,.0f}. NASENI levy is ₦{t['naseni']:,.0f}. Current tax expense is ₦{t['current']:,.0f}.",
            f"The current-tax liability on the statement of financial position is ₦{s['tax_cr']:,.0f}. It is the charge of each year less tax payments the narrations identify (₦{model['bank'].loc[model['bank']['category']=='TAX_PAID','dr'].sum():,.0f} across the three years). Many tax payments may not say 'tax' in the narration and will be inside payments pending allocation. The liability should fall if those payments are identified. It should not be read as an agreed assessment.",
            "Deferred tax is not recognised. Temporary differences cannot be measured reliably without a capital-allowance register, and a deferred-tax asset on the loss allowance is not material.",
        ]),
        ("13. Share capital and opening balances", [
            "Recognised share capital is ₦1,000,000. The CAC status report of 26 August 2026 records ₦100,000,000 issued. The ₦99,000,000 difference was not traced to these bank accounts in 2023–2025 and is not recognised. The board should produce the return of allotment and evidence of consideration.",
            "At 1 January 2023 the only opening balance taken from the statements is cash of ₦478,080.25 (Providus ₦464,164.15, First Bank ₦12,971.91, Sterling ₦944.19, GTBank nil). Equity equals that cash. There is no opening inventory, receivable or payable, because none could be reconstructed. Land sold in 2023–2025 may have been bought before 2023. If it was, cost of sales is understated and profit is overstated. That is a qualification, not a hidden assumption.",
        ]),
        ("14. Commitments, exceptions and items held out of turnover", [
            f"The exceptions tab holds ₦{model['exceptions_amt']:,.0f} of secondary records for December 2023 to December 2024 that were not added to the sales ledger, because most are aggregates or splits of receipts already counted. They are not in revenue and they are not in deferred income. Adding them would double-count. The auditor should still agree any line that is not an aggregate.",
            "Nine client-register entries dated 2026 were left out of the subscription population. They are outside the period.",
            "No dividend is proposed beyond the distributions already traced. There is no capital commitment recorded in the files other than the housing and land pipeline.",
        ]),
        ("15. Comparison with the pack that was circulated and rejected", [
            _compare_table(model, year, old),
            "The rejected pack's revenue is not a subset of these collections. It does not reconcile to the sales ledger at all. The cash figure does not reconcile to the bank. These revised statements do both, and they show the reconciling items instead of removing them.",
        ]),
        ("16. Going concern and events after the reporting period", [
            f"At 31 December 2025 the company held ₦155.9 million of cash, deferred customer collections of ₦1.22 billion, and a housing and land pipeline. The files do not show a borrowing covenant or a demand that would stop the business within twelve months of the date of these statements. Going concern is adopted. It is qualified by the unallocated payments and by the related-party debit, both of which the board must explain.",
            "The bank extraction and the sales ledger stop at 31 December 2025 for this revision. Transactions in 2026, including the construction engagement noted in the source file, are not in these statements.",
        ]),
    ]


def _cash_table(model, year):
    rows = [["Bank", "Account", "Name on statement", f"31 Dec {year}"]]
    names = {
        "GTBank": ("0893079582", "BAAY PROJECTS LTD"),
        "First Bank": ("2033736938", "BAAY DEGOK NIG LTD"),
        "Sterling": ("0091166190", "BAAY PROJECTS"),
        "Providus": ("5400724970", "BAAY PROJECTS LTD"),
    }
    total = 0
    for bk, (acct, name) in names.items():
        bal = model["stated_cash"][year]["by_bank"][bk]
        total += bal
        rows.append([bk, acct, name, N(bal, dash_zero=False)])
    rows.append(["Total cash", "", "", N(total, dash_zero=False)])
    return tbl(rows, [80, 90, 230, 123], total_rows={5})


def _compare_table(model, year, old):
    p = model["pnl"][year]
    s = model["sfp"][year]
    rows = [
        ["", "Previously circulated", "Revised", "What changed"],
        ["Turnover — collections", N(old["revenue"]), N(model["collections"][year]), "The sales ledger, in full"],
        ["Revenue recognised", "—", N(p["revenue"]), "Only the earned portion"],
        ["Deferred income", N(old["deferred"]), N(s["contract_liabilities"]), "Unearned collections"],
        ["Receivables", N(old["receivables"]), N(s["receivables_net"]), "Earned amounts still unpaid"],
        ["Cash", N(old["cash"]), N(s["cash"]), "The four bank accounts"],
        ["Total assets", N(old["assets"]), N(s["total_assets"]), "Includes items still to allocate"],
        ["Profit before tax", N(old["pbt"]), N(p["pbt"]), "Before unidentified payments"],
    ]
    return tbl(rows, [120, 115, 115, 173], text_cols={3})


def build_basis_report(model, path):
    story = []
    story.append(Spacer(1, 28))
    story.append(P("REVISED", "CoverCo"))
    story.append(P("BAAY PROJECTS LIMITED", "CoverTitle"))
    story.append(P("RC 1526224", "CoverSub"))
    story.append(Spacer(1, 6))
    story.append(hr())
    story.append(P("BASIS OF REVISION", "CoverTitle"))
    story.append(P("How the 2023, 2024 and 2025 financial statements were rebuilt from the actual transactions", "CoverSub"))
    story.append(Spacer(1, 8))
    story.append(P(
        "This report explains the revision requested because the circulated audited financial statements "
        "did not reflect the company's turnover, receivables, deferred income, assets and liabilities. "
        "It is the document to read before the three annual packs. All amounts are in naira.",
        "CoverSub"))
    story.append(Spacer(1, 8))
    story.append(tbl([
        ["", "2023", "2024", "2025"],
        ["Turnover in the rejected pack", "38,500,000", "64,200,000", "112,500,000"],
        ["Turnover in the sales ledger", N(model["collections"][2023]), N(model["collections"][2024]), N(model["collections"][2025])],
        ["Cash in the rejected pack", "2,842,403", "5,142,678", "8,924,048"],
        ["Cash on the bank statements", N(model["sfp"][2023]["cash"]), N(model["sfp"][2024]["cash"]), N(model["sfp"][2025]["cash"])],
        ["Deferred income, revised", N(model["sfp"][2023]["contract_liabilities"]), N(model["sfp"][2024]["contract_liabilities"]), N(model["sfp"][2025]["contract_liabilities"])],
    ], [180, 114, 114, 115], bold_rows={2, 4}))
    story.append(Spacer(1, 8))

    story.append(P("1. What was wrong with the circulated pack", "H1"))
    story.append(hr())
    story.append(P(
        "The circulated statements, including the version marked updated, still reported revenue of "
        "₦38.5 million, ₦64.2 million and ₦112.5 million. The sales ledger on the same branch totals "
        "₦125.6 million, ₦594.1 million and ₦1,865.1 million. The circulated cash balances were about "
        "₦2.8 million, ₦5.1 million and ₦8.9 million. The bank statements show ₦12.9 million, ₦12.9 million "
        "and ₦155.9 million. The circulated pack also kept a 2022 cash figure of ₦52,500 when the statements "
        "show ₦478,080, and it cited bank account numbers that are not the extracted accounts. Deferred income "
        "was shown at ₦2.4 million to ₦5.8 million against housing collections, on the client register, that "
        "had not been handed over. Those statements could not be agreed to the company's transactions. They "
        "have been replaced, not patched.",
        "Body"))

    story.append(P("2. Three-year position after the revision", "H1"))
    story.append(tbl([
        ["", "2023", "2024", "2025"],
        ["Customer collections (turnover)", N(model["collections"][2023]), N(model["collections"][2024]), N(model["collections"][2025])],
        ["Revenue recognised", N(model["pnl"][2023]["revenue"]), N(model["pnl"][2024]["revenue"]), N(model["pnl"][2025]["revenue"])],
        ["   Land", N(model["pnl"][2023]["rev_land"]), N(model["pnl"][2024]["rev_land"]), N(model["pnl"][2025]["rev_land"])],
        ["   Housing", N(model["pnl"][2023]["rev_hous"]), N(model["pnl"][2024]["rev_hous"]), N(model["pnl"][2025]["rev_hous"])],
        ["Deferred income", N(model["sfp"][2023]["contract_liabilities"]), N(model["sfp"][2024]["contract_liabilities"]), N(model["sfp"][2025]["contract_liabilities"])],
        ["Trade receivables, net", N(model["sfp"][2023]["receivables_net"]), N(model["sfp"][2024]["receivables_net"]), N(model["sfp"][2025]["receivables_net"])],
        ["Cash at bank", N(model["sfp"][2023]["cash"]), N(model["sfp"][2024]["cash"]), N(model["sfp"][2025]["cash"])],
        ["Inventory", N(model["sfp"][2023]["inventory"]), N(model["sfp"][2024]["inventory"]), N(model["sfp"][2025]["inventory"])],
        ["Co-ownership liability", N(model["sfp"][2023]["co_ownership"]), N(model["sfp"][2024]["co_ownership"]), N(model["sfp"][2025]["co_ownership"])],
        ["Receipts pending allocation", N(model["sfp"][2023]["receipts_pending"]), N(model["sfp"][2024]["receipts_pending"]), N(model["sfp"][2025]["receipts_pending"])],
        ["Payments pending allocation", N(model["sfp"][2023]["payments_pending"]), N(model["sfp"][2024]["payments_pending"]), N(model["sfp"][2025]["payments_pending"])],
        ["Profit before tax (before that allocation)", N(model["pnl"][2023]["pbt"]), N(model["pnl"][2024]["pbt"]), N(model["pnl"][2025]["pbt"])],
        ["Total assets", N(model["sfp"][2023]["total_assets"]), N(model["sfp"][2024]["total_assets"]), N(model["sfp"][2025]["total_assets"])],
    ], [200, 107, 108, 108], bold_rows={1, 2, 5}, total_rows={13}))
    story.append(Spacer(1, 6))
    story.append(P(
        "Each statement of financial position balances. Cash equals the bank statements. Revenue equals "
        "sales-ledger collections plus the movement in receivables minus the movement in deferred income. "
        "Those three ties are the control. They replace the previous pack's internal tie, which balanced "
        "only because both sides used the same unsupported figures.",
        "Body"))

    story.append(P("3. How turnover, receivables and deferred income were set", "H1"))
    story.append(P(
        "The sales ledger is the turnover population: 859 lines, ₦2,584,814,000, land and housing, no construction "
        "in the period. It was not scaled and it was not replaced with a percentage of bank inflows.",
        "Body"))
    story.append(P(
        "Land subscriptions marked bought or closed start from the contract value. The unpaid balance is a "
        "receivable only up to the balance on the client register. Where the register is lower, revenue is "
        "reduced by the same amount. That reduction is ₦515,131,000 across the three years. It is disclosed "
        "so it can be tested. It is not booked as a receivable the register does not support. Housing is recognised only "
        "on evidenced handover (Bought, Completed, or Paid with the register showing the money). Everything else "
        "the customer has paid stays in deferred income. That is why 2025 housing revenue is ₦814.6 million "
        "while housing collections across the three years are ₦1.90 billion, and why deferred income at "
        "31 December 2025 is ₦1.22 billion. The collections are in the statements. They are in the liability "
        "until handover, which is what IFRS 15 requires and what the client register supports.",
        "Body"))
    story.append(P(
        "Receivables are deliberately small. The register's large unpaid housing balances belong to contracts "
        "that have not been handed over. Recognising them as receivables would also require recognising the "
        "revenue, which the evidence does not support. They are disclosed as a pipeline in Note 6 of each "
        "annual pack, so the amount still to be collected is visible without being mislabelled.",
        "Body"))

    story.append(P("4. How the bank was integrated", "H1"))
    story.append(P(
        "All 23,173 transactions on the four accounts for 1 January 2023 to 31 December 2025 were classified. "
        "Gross inflows are ₦4.11 billion and gross outflows are ₦3.95 billion. A large part of both is movement "
        "between the company's own accounts and is eliminated. What remains is split into customer property "
        "receipts, trading income, co-ownership, director movements, inventory purchases, operating costs, tax, "
        "and two pending-allocation accounts.",
        "Body"))
    story.append(P(
        "Identified property receipts in the banks are ₦2.08 billion. The sales ledger is ₦2.58 billion. The "
        "cumulative gap of ₦506.4 million at 31 December 2025 is an asset: customer receipts recorded by client "
        "relations and not covered by property receipts identified in these four accounts. It is not extra cash, "
        "and it is not ignored. Line-by-line matching agreed ₦1.45 billion to a specific credit. The statements "
        "use the cumulative gap rather than the larger line-level miss, so that a receipt which is in the bank "
        "but not matched to a single line is not counted twice.",
        "Body"))
    story.append(P(
        "Co-ownership of ₦235.7 million was taken out of turnover by the source file and is a liability, net of "
        "returns. Credit memos of ₦613,250 are non-cash and are not posted. Exceptions of ₦138.4 million were "
        "left out of the sales ledger by the source file to avoid double counting, and they stay out.",
        "Body"))

    story.append(P("5. What is still open — and why profit is qualified", "H1"))
    story.append(P(
        "Payments pending allocation at 31 December 2025 are ₦1,055 million. Receipts pending allocation are "
        "₦873 million. Net transfers with company accounts that do not mirror inside the four extracted accounts "
        "are ₦319 million. Amounts due from related parties are ₦137 million. Loan repayments described as such "
        "in the narrations exceed loan proceeds traced, by ₦68 million. Tax identified as paid is ₦8.6 million "
        "against a computed liability of ₦184 million; further tax payments are likely inside the pending payments.",
        "Body"))
    story.append(P(
        "Profit before tax of ₦25.9 million, ₦151.5 million and ₦404.8 million is profit before those pending "
        "payments are allocated. It is also profit before any land cost incurred before 1 January 2023, which "
        "is not in the files. The draft audit opinion is qualified on both points. The profit figure is in the "
        "statements because the identified transactions produce it. It is not a finished result, and the "
        "directors' report says so in terms.",
        "Body"))

    story.append(P("6. Files labelled revised", "H1"))
    story.append(tbl([
        ["File", "What the revision added"],
        ["Bank extraction, revised", "Every transaction classified. Cash proof by bank and year. Unallocated items of ₦1 million and above listed with the narration."],
        ["Sales and customers, revised", "The original population is unchanged. Recognition, the receivable and deferred income are added on a customer line."],
        ["Sales-to-bank reconciliation, revised", "Bridge from the sales ledger to the bank extraction. The previous bridge used bank totals that are not in the statements."],
        ["Master workbook, revised", "The statements, the tax computation, the contract schedule, and the checks that cash, turnover and the balance sheet tie."],
        ["Annual statements, 2023, 2024 and 2025", "One pack a year, with this basis report. The auditor's report in each pack is marked as a draft."],
    ], [170, 353], text_cols={0, 1}))
    story.append(Spacer(1, 8))
    story.append(P(
        "Checks performed and passed: assets equal liabilities and equity in each year; cash equals the stated "
        "bank balances; the sales-ledger total is ₦2,584,814,000; revenue equals collections plus the receivable "
        "movement minus the deferred-income movement; no amount was posted to revenue as a plug.",
        "Body"))
    story.append(Spacer(1, 10))
    story.append(P(
        "Prepared 8 October 2026 for the board of Baay Projects Limited and for Sanni Waheed &amp; Co. "
        "The annual packs carry the statements. This report carries the reason they were rebuilt.",
        "Small"))

    doc = SimpleDocTemplate(
        path, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=40, bottomMargin=36,
        title="Baay Projects Limited — Basis of revision 2023-2025",
        author="Revised from company records",
    )
    doc.build(story, canvasmaker=RevCanvas)
    print("wrote", path)


def main():
    model = pickle.load(open("/tmp/baay_extract/model.pkl", "rb"))
    build_basis_report(model, "BAAY_PROJECTS_LIMITED_Revision_Basis_Report_revised.pdf")
    for year in (2023, 2024, 2025):
        build_afs(model, year, f"BAAY_PROJECTS_LIMITED_AFS_{year}_revised.pdf")


if __name__ == "__main__":
    main()
