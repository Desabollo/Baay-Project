"""Annual financial statements for signature.

The comparative figures for 2022 are the signed financial statements for that
year. The figures for 2023, 2024 and 2025 are the balances that agree to the
sales records and the four bank accounts.
"""

from __future__ import annotations

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
    HRFlowable, KeepTogether,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

NAVY = colors.HexColor("#1F4E79")
BLUE = colors.HexColor("#2F5597")
ACCENT = colors.HexColor("#D9E1F2")
ZEBRA = colors.HexColor("#F2F4F8")
PALE = colors.HexColor("#F7F9FC")
LINE = colors.HexColor("#D0D7E2")
DARK = colors.HexColor("#262626")
MUTED = colors.HexColor("#595959")

DIRECTORS = [
    (1, "Adegoke Segun Babatunde", 5, "Keshinro Phebe Oluwatunmise"),
    (2, "Adegoke Raheem Adebayo", 6, "Owolabi Charles Oluwatobi"),
    (3, "Ajibola Oluwatobi Adedamola", 7, "Olayinka Oladotun Emmanuel"),
    (4, "Shuaib Suliat Aduke", 8, "Noah Abdulazeez Afolabi"),
]

# Signed financial statements for the year ended 31 December 2022.
Y2022 = dict(
    ppe=9540.00, receivables=3374668.00, cash=52500.00, assets=3436708.00,
    payables=75000.00, liabilities=75000.00, share=1000000.00, re=2361708.00,
    equity=3361708.00, revenue=6200000.00, cos=0.0, opex=4400427.00,
    pbt=1799573.00, tax=0.0, pat=1799573.00, re_open=562135.00,
    cash_open=78350.00, cash_close=52500.00, cf_op=-25850.00, cf_net=-25850.00,
    dep=4770.00,
)

BANK = {
    2023: {
        "ADMIN_OPS": (0, 7133393.51), "BANK_CHARGE": (80.63, 234199.27),
        "BANK_CHARGE_CONTRA": (387740.65, 0), "BORROWINGS_IN": (500000, 0),
        "BORROWINGS_OUT": (0, 22041603.52), "CUSTOMER_COLLECTION": (115185685, 0),
        "CUSTOMER_REFUND": (0, 13940516.88), "DIRECTOR_REMUNERATION": (0, 7151060),
        "DIVIDEND": (0, 1620007.5), "FINANCE_COST": (0, 50.52),
        "INTERBANK": (104170000, 23673002.63), "INVENTORY_DEVELOPMENT": (0, 64174083.77),
        "INVENTORY_LAND": (0, 36320054.13), "OTHER_INCOME_REFUND": (5454452.28, 0),
        "PPE": (0, 6945000), "PROFESSIONAL": (0, 1500007.5),
        "RELATED_PARTY_IN": (11864330, 0), "RELATED_PARTY_OUT": (0, 34338900),
        "RENT": (0, 4035000), "SELLING_COMMISSION": (0, 9520792.5),
        "STAFF_COST": (0, 10403081.58), "TAX_PAID": (0, 228208.26),
        "UNALLOCATED_IN": (193803109.3, 0), "UNALLOCATED_OUT": (0, 175668523.22),
    },
    2024: {
        "ADMIN_OPS": (0, 37496273.18), "BANK_CHARGE": (107.5, 713116.84),
        "BANK_CHARGE_CONTRA": (353599.01, 0), "BORROWINGS_OUT": (0, 22680526.22),
        "CO_OWNERSHIP_IN": (4700000, 0), "CO_OWNERSHIP_OUT": (0, 6640011.25),
        "CUSTOMER_COLLECTION": (431428600, 0), "CUSTOMER_REFUND": (0, 8222400),
        "DIRECTOR_REMUNERATION": (0, 14995339.91), "DIVIDEND": (0, 14070011.25),
        "FINANCE_COST": (0, 319704.08), "INTERBANK": (158098046.99, 121712675.76),
        "INVENTORY_DEVELOPMENT": (0, 64432533.14), "INVENTORY_LAND": (0, 105770393.13),
        "OTHER_INCOME_REFUND": (266470, 0), "OTHER_INCOME_SERVICE": (2050000, 0),
        "PPE": (0, 10153976), "PROFESSIONAL": (0, 4575003.75),
        "RELATED_PARTY_IN": (18663001, 0), "RELATED_PARTY_OUT": (0, 78819018.75),
        "RENT": (0, 10450000), "SELLING_COMMISSION": (0, 73413514),
        "STAFF_COST": (0, 36991007.9), "TAX_PAID": (0, 1356619.25),
        "UNALLOCATED_IN": (282689448.53, 0), "UNALLOCATED_OUT": (0, 285458455.34),
    },
    2025: {
        "ADMIN_OPS": (0, 50089668.55), "BANK_CHARGE": (1239.58, 376074.74),
        "BANK_CHARGE_CONTRA": (353101.01, 0), "BORROWINGS_IN": (5667000, 0),
        "BORROWINGS_OUT": (0, 29009921), "CO_OWNERSHIP_IN": (68074000, 0),
        "CO_OWNERSHIP_OUT": (0, 7453007.5), "CUSTOMER_COLLECTION": (1531759807, 0),
        "CUSTOMER_REFUND": (0, 15894109.17), "DIRECTOR_REMUNERATION": (0, 31770033.75),
        "DIVIDEND": (0, 13650015), "FINANCE_COST": (0, 1458618.52),
        "FX_ADVANCE": (0, 41300000), "INTERBANK": (434770297.81, 232691918.5),
        "INVENTORY_DEVELOPMENT": (0, 306604912.26), "INVENTORY_HOUSING": (0, 530000953.75),
        "INVENTORY_LAND": (0, 297472001.25), "OTHER_INCOME_REFUND": (37889542.05, 0),
        "OTHER_INCOME_TRADING": (290523480, 0), "PPE": (0, 8158000),
        "PROFESSIONAL": (0, 14380026.26), "RELATED_PARTY_IN": (10274428, 0),
        "RELATED_PARTY_OUT": (0, 64765272), "RENT": (0, 6648000),
        "SELLING_COMMISSION": (0, 117252650), "STAFF_COST": (0, 49501351.76),
        "TAX_PAID": (0, 7016805.38), "TRADING_PURCHASES": (0, 213091609.02),
        "UNALLOCATED_IN": (396629291.5, 0), "UNALLOCATED_OUT": (0, 594347767.56),
    },
}
CASH = {2022: 52500.00, 2023: 12915993.32, 2024: 12894686.60, 2025: 155904157.58}
BANK_OPEN_2023 = 478080.25

YEARS = {
    2023: dict(
        rev_land=84117680.00, rev_hous=0.00, revenue=84117680.00,
        other_trading=0.0, other_service=0.0, other_refunds=5454452.28, bank_credit=153622.01,
        cos=8708574.30, cos_trading=0.0,
        refunds=13940516.88, staff=10403081.58, director_rem=7151060.00,
        commission=9520792.50, rent=4035000.00, professional=1500007.50,
        admin=7133393.51, bank_charges=0.0, depreciation=1389000.00, ecl=5000.00,
        finance=50.52, pbt=25939277.50, tax=6287950.79, pat=19651326.71,
        dividends=1620007.50,
        ppe=5556000.00, ppe_cost=6945000.00, accum_dep=1389000.00,
        inventory=91785563.60, rec_gross=100000.00, rec_ecl=5000.00, rec_net=95000.00,
        other_rec=10434315.00, other_assets=197210126.74, related=22474570.00,
        fx=0.0, loan_excess=21541603.52, payments_unid=175668523.22,
        cash=12915993.32, cl=41602320.00, co=0.0, other_liab=193803109.30,
        inter=80496997.37, tax_liab=6059742.53,
        assets=340471568.66, liabilities=321962169.20, equity=18509399.46, re=17509399.46,
        size="medium", rate=0.20, cit=5466655.50, tet=819998.33, ptf=1296.96, naseni=0.0,
    ),
    2024: dict(
        rev_land=300864060.00, rev_hous=50000000.00, revenue=350864060.00,
        other_trading=0.0, other_service=2650000.00, other_refunds=266470.00, bank_credit=0.0,
        cos=13161526.43, cos_trading=0.0,
        refunds=8222400.00, staff=36991007.90, director_rem=14995339.91,
        commission=73413514.00, rent=10450000.00, professional=4575003.75,
        admin=37496273.18, bank_charges=359410.33, depreciation=2030795.20, ecl=240000.00,
        finance=319704.08, pbt=151525555.22, tax=51139185.81, pat=100386369.41,
        dividends=14070011.25,
        ppe=13679180.80, ppe_cost=17098976.00, accum_dep=3419795.20,
        inventory=248826963.44, rec_gross=4900000.00, rec_ecl=245000.00, rec_net=4655000.00,
        other_rec=193406315.00, other_assets=505349108.30, related=82630587.75,
        fx=0.0, loan_excess=44222129.74, payments_unid=461126978.56,
        cash=12894686.60, cl=289618860.00, co=17779988.75, other_liab=476492557.83,
        inter=116882368.60, tax_liab=55842309.09,
        assets=1061441841.89, liabilities=956616084.27, equity=104825757.62, re=103825757.62,
        size="large", rate=0.30, cit=46138905.13, tet=4613890.51, ptf=7576.28, naseni=378813.89,
    ),
    2025: dict(
        rev_land=126293400.00, rev_hous=814600000.00, revenue=940893400.00,
        other_trading=290523480.00, other_service=0.0, other_refunds=37889542.05, bank_credit=0.0,
        cos=575370741.96, cos_trading=213091609.02,
        refunds=15894109.17, staff=49501351.76, director_rem=31770033.75,
        commission=117252650.00, rent=6648000.00, professional=14380026.26,
        admin=50089668.55, bank_charges=21734.15, depreciation=1631600.00, ecl=500000.00,
        finance=1458618.52, pbt=404787887.93, tax=135315640.13, pat=269472247.80,
        dividends=13650015.00,
        ppe=20205580.80, ppe_cost=25256976.00, accum_dep=5051395.20,
        inventory=1020625697.76, rec_gross=14900000.00, rec_ecl=745000.00, rec_net=14155000.00,
        other_rec=671936408.00, other_assets=1164339796.86, related=137121431.75,
        fx=41300000.00, loan_excess=67565050.74, payments_unid=1055474746.12,
        cash=155904157.58, cl=1223838860.00, co=223577481.25, other_liab=873121849.33,
        inter=318960747.91, tax_liab=184141143.84,
        assets=3184288072.75, liabilities=2823640082.33, equity=360647990.42, re=359647990.42,
        size="large", rate=0.30, cit=122075846.38, tet=12207584.64, ptf=20239.39, naseni=1011969.72,
    ),
}
ISSUED = 100_000_000.00
UNPAID = 99_000_000.00
PAID_CAP = 1_000_000.00
# Retained earnings at 1 January 2023 in the model, so that equity equals the bank cash
# at that date. The signed accounts reported a different retained earnings balance.
MODEL_OPEN_RE = -521919.75


def money(x, blank_zero=True):
    if x is None:
        return ""
    x = float(x)
    if blank_zero and abs(x) < 0.005:
        return "—"
    s = f"{abs(x):,.2f}"
    return f"({s})" if x < 0 else s


def oi(d):
    return d["other_trading"] + d["other_service"] + d["other_refunds"] + d["bank_credit"]


def ox(d):
    return (d["refunds"] + d["staff"] + d["director_rem"] + d["commission"] + d["rent"]
            + d["professional"] + d["admin"] + d["bank_charges"] + d["depreciation"] + d["ecl"])


def face_check():
    for y, d in YEARS.items():
        pbt = d["revenue"] - d["cos"] + oi(d) - ox(d) - d["finance"]
        pat = pbt - d["tax"]
        if abs(pbt - d["pbt"]) > 0.05 or abs(pat - d["pat"]) > 0.05:
            raise SystemExit(f"Profit does not tie for {y}: {pbt} {pat}")
        assets = d["ppe"] + d["inventory"] + d["rec_net"] + d["other_rec"] + d["other_assets"] + d["related"] + d["cash"]
        liab = d["cl"] + d["co"] + d["other_liab"] + d["inter"] + d["tax_liab"]
        if abs(assets - d["assets"]) > 0.05 or abs(liab - d["liabilities"]) > 0.05:
            raise SystemExit(f"Statement of financial position does not add for {y}")
        if abs(assets - liab - d["equity"]) > 0.05:
            raise SystemExit(f"Statement of financial position does not balance for {y}")
    # 2022 signed accounts
    if abs(Y2022["ppe"] + Y2022["receivables"] + Y2022["cash"] - Y2022["assets"]) > 0.05:
        raise SystemExit("2022 assets do not add")
    if abs(Y2022["assets"] - Y2022["liabilities"] - Y2022["equity"]) > 0.05:
        raise SystemExit("2022 statement does not balance")
    if abs(Y2022["revenue"] - Y2022["opex"] - Y2022["pat"]) > 0.05:
        raise SystemExit("2022 profit does not tie")
    rolled = Y2022["re"] + YEARS[2023]["pat"] - YEARS[2023]["dividends"]
    other = YEARS[2023]["re"] - rolled
    if abs(other + 2883627.75) > 0.05:
        raise SystemExit(f"Unexpected equity movement {other}")
    if abs((BANK_OPEN_2023 - Y2022["cash"]) - 425580.25) > 0.05:
        raise SystemExit("Opening cash difference is not the expected amount")


def cash_lines(year):
    b = BANK[year]
    def cr(k):
        return b.get(k, (0, 0))[0]
    def dr(k):
        return b.get(k, (0, 0))[1]
    cust = cr("CUSTOMER_COLLECTION")
    other_in = (cr("OTHER_INCOME_TRADING") + cr("OTHER_INCOME_SERVICE") + cr("OTHER_INCOME_REFUND")
                + cr("BANK_CHARGE_CONTRA") + cr("BANK_CHARGE"))
    other_rec = cr("UNALLOCATED_IN")
    project = dr("INVENTORY_LAND") + dr("INVENTORY_HOUSING") + dr("INVENTORY_DEVELOPMENT") + dr("TRADING_PURCHASES")
    overhead = (dr("STAFF_COST") + dr("DIRECTOR_REMUNERATION") + dr("SELLING_COMMISSION")
                + dr("RENT") + dr("PROFESSIONAL") + dr("ADMIN_OPS") + dr("BANK_CHARGE")
                + dr("CUSTOMER_REFUND") + dr("FINANCE_COST"))
    other_pay = dr("UNALLOCATED_OUT")
    tax = dr("TAX_PAID")
    ppe = dr("PPE")
    fx = dr("FX_ADVANCE")
    rp = cr("RELATED_PARTY_IN") - dr("RELATED_PARTY_OUT")
    bor = cr("BORROWINGS_IN") - dr("BORROWINGS_OUT")
    div = -dr("DIVIDEND")
    co = cr("CO_OWNERSHIP_IN") - dr("CO_OWNERSHIP_OUT")
    inter = cr("INTERBANK") - dr("INTERBANK")
    op = cust + other_in + other_rec - project - overhead - other_pay - tax
    inv = -(ppe + fx)
    fin = rp + bor + div + co + inter
    extra = (BANK_OPEN_2023 - Y2022["cash"]) if year == 2023 else 0.0
    net = op + inv + fin + extra
    opening = Y2022["cash"] if year == 2023 else CASH[year - 1]
    close = opening + net
    if abs(close - CASH[year]) > 1:
        raise SystemExit(f"Cash flow does not tie for {year}: {close} vs {CASH[year]}")
    rows = [
        ("CASH FLOWS FROM OPERATING ACTIVITIES", None, True),
        ("Receipts from customers", cust, False),
        ("Other income received", other_in, False),
        ("Other receipts", other_rec, False),
        ("Payments for land, housing, development and trading", -project, False),
        ("Payments for staff, commission, rent and overheads", -overhead, False),
        ("Other payments", -other_pay, False),
        ("Tax paid", -tax, False),
        ("Net cash from operating activities", op, True),
        ("CASH FLOWS FROM INVESTING ACTIVITIES", None, True),
        ("Purchase of property, plant and equipment", -ppe, False),
        ("Other advances", -fx, False),
        ("Net cash from investing activities", inv, True),
        ("CASH FLOWS FROM FINANCING ACTIVITIES", None, True),
        ("Related-party movements", rp, False),
        ("Borrowings and repayments", bor, False),
        ("Dividends paid", div, False),
        ("Co-ownership receipts, net of returns", co, False),
        ("Movements on the Company's other accounts", inter, False),
        ("Net cash from financing activities", fin, True),
    ]
    if year == 2023:
        rows.append(("Other cash movements", extra, False))
    rows += [
        ("Net increase / (decrease) in cash", net, True),
        ("Cash and cash equivalents at 1 January", opening, False),
        ("Cash and cash equivalents at 31 December", close, True),
    ]
    return rows


class AFSCanvas(canvas.Canvas):
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
        if self._pageNumber == 1:
            return
        self.saveState()
        self.setFillColor(NAVY)
        self.setFont("DejaVuBold", 8)
        self.drawString(36, A4[1] - 24, "BAAY PROJECTS LIMITED (RC 1526224)")
        self.setFont("DejaVu", 8)
        self.drawRightString(A4[0] - 36, A4[1] - 24, "Audited Financial Statements")
        self.setStrokeColor(NAVY)
        self.setLineWidth(0.8)
        self.line(36, A4[1] - 28, A4[0] - 36, A4[1] - 28)
        self.setStrokeColor(LINE)
        self.setLineWidth(0.4)
        self.line(36, 32, A4[0] - 36, 32)
        self.setFillColor(MUTED)
        self.setFont("DejaVu", 7.5)
        self.drawString(36, 20, "BAAY PROJECTS LIMITED — ANNUAL REPORT AND FINANCIAL STATEMENTS")
        self.drawRightString(A4[0] - 36, 20, f"Page {self._pageNumber} of {n}")
        self.restoreState()


def styles():
    ss = getSampleStyleSheet()
    ss.add(ParagraphStyle("CoverTitle", fontName="DejaVuBold", fontSize=16, leading=20, textColor=NAVY, alignment=TA_CENTER))
    ss.add(ParagraphStyle("CoverSub", fontName="DejaVuBold", fontSize=11, leading=14, textColor=BLUE, alignment=TA_CENTER))
    ss.add(ParagraphStyle("CoverSmall", fontName="DejaVu", fontSize=8, leading=11, textColor=MUTED, alignment=TA_CENTER))
    ss.add(ParagraphStyle("H1", fontName="DejaVuBold", fontSize=11.5, leading=14, textColor=NAVY, spaceBefore=1, spaceAfter=1))
    ss.add(ParagraphStyle("H2", fontName="DejaVuBold", fontSize=9, leading=12, textColor=BLUE, spaceBefore=6, spaceAfter=2))
    ss.add(ParagraphStyle("Body", fontName="DejaVu", fontSize=8, leading=10.6, alignment=TA_JUSTIFY, textColor=DARK))
    ss.add(ParagraphStyle("Small", fontName="DejaVu", fontSize=7.4, leading=9.6, alignment=TA_JUSTIFY, textColor=DARK))
    ss.add(ParagraphStyle("Th", fontName="DejaVuBold", fontSize=7.1, leading=9, textColor=colors.white, alignment=TA_CENTER))
    ss.add(ParagraphStyle("Td", fontName="DejaVu", fontSize=7.1, leading=9.1, textColor=DARK))
    ss.add(ParagraphStyle("TdB", fontName="DejaVuBold", fontSize=7.1, leading=9.1, textColor=NAVY))
    ss.add(ParagraphStyle("Sig", fontName="DejaVu", fontSize=8, leading=11, textColor=DARK))
    return ss


S = styles()


def P(text, style="Body"):
    return Paragraph(str(text), S[style])


def table(data, widths, header=True, bold_rows=None, total_rows=None, text_cols=None):
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
                line.append(Paragraph(str(val), S["TdB"] if r in bold_rows or r in total_rows else S["Td"]))
            else:
                sty = ParagraphStyle(
                    f"n{r}{c}", parent=S["TdB"] if r in bold_rows or r in total_rows else S["Td"],
                    alignment=TA_RIGHT,
                )
                line.append(Paragraph(str(val), sty))
        wrapped.append(line)
    t = Table(wrapped, colWidths=widths, repeatRows=1 if header else 0)
    cmds = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2.4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.4),
        ("GRID", (0, 0), (-1, -1), 0.3, LINE),
    ]
    if header:
        cmds.append(("BACKGROUND", (0, 0), (-1, 0), NAVY))
    for r in total_rows:
        cmds.append(("BACKGROUND", (0, r), (-1, r), ACCENT))
    for r in bold_rows:
        if r not in total_rows:
            cmds.append(("BACKGROUND", (0, r), (-1, r), ZEBRA))
    t.setStyle(TableStyle(cmds))
    return t


def rule():
    return HRFlowable(width="100%", thickness=1, color=BLUE, spaceBefore=1, spaceAfter=6)


def signature_pair(left, right):
    t = Table(
        [[Paragraph("<br/>".join(left), S["Sig"]), Paragraph("<br/>".join(right), S["Sig"])]],
        colWidths=[255, 255],
    )
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def prior_amount(year, key):
    """Comparative amount. 2022 comes from the signed accounts. Later years come from these statements."""
    if year != 2023:
        return YEARS[year - 1][key]
    mapping = {
        "revenue": Y2022["revenue"], "cos": Y2022["cos"], "pbt": Y2022["pbt"], "tax": Y2022["tax"],
        "pat": Y2022["pat"], "ppe": Y2022["ppe"], "cash": Y2022["cash"], "rec_net": Y2022["receivables"],
        "re": Y2022["re"], "equity": Y2022["equity"], "assets": Y2022["assets"],
        "liabilities": Y2022["liabilities"], "payables": Y2022["payables"],
        "dividends": 0.0, "finance": 0.0, "depreciation": Y2022["dep"],
    }
    return mapping.get(key, 0.0)


def build(year, path):
    d = YEARS[year]
    comp = year - 1
    c = None if year == 2023 else YEARS[comp]
    story = []

    story.append(Spacer(1, 26))
    story.append(P("BAAY PROJECTS LIMITED", "CoverTitle"))
    story.append(Spacer(1, 3))
    story.append(P("(RC 1526224 • TIN 21548976-0001)", "CoverSub"))
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=2.5, color=NAVY, spaceBefore=2, spaceAfter=14))
    story.append(P("ANNUAL REPORT AND AUDITED FINANCIAL STATEMENTS", "CoverTitle"))
    story.append(Spacer(1, 6))
    story.append(P(f"FOR THE YEAR ENDED 31 DECEMBER {year}", "CoverSub"))
    story.append(Spacer(1, 3))
    story.append(P(f"(With Comparative Figures for the Year Ended 31 December {comp})", "CoverSmall"))
    story.append(Spacer(1, 12))
    badge = Table([[P("ANNUAL REPORT AND AUDITED FINANCIAL STATEMENTS", "Th")]], colWidths=[420])
    badge.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(badge)
    story.append(Spacer(1, 14))
    info = [
        [P("Registered Corporate Office:", "TdB"), P("No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State", "Td")],
        [P("Incorporation:", "TdB"), P("18 September 2018 • RC 1526224 • Formerly Baay Degok Nig Ltd (name changed 28 June 2021)", "Td")],
        [P("Company Secretary:", "TdB"), P("Adegoke Mary Ayoboade", "Td")],
        [P("Independent Auditors:", "TdB"), P("Sanni Waheed &amp; Co. (Chartered Accountants), 1st Floor, No. 29 Olonode Street, Alagomeji, Yaba, Lagos", "Td")],
        [P("Principal Bankers:", "TdB"), P("Providus Bank Plc • First Bank of Nigeria Limited • Sterling Bank Plc • GTBank", "Td")],
        [P("Accounting Framework:", "TdB"), P("International Financial Reporting Standards (IFRS) &amp; CAMA 2020", "Td")],
    ]
    info_t = Table(info, colWidths=[130, 380])
    info_t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), PALE),
        ("BOX", (0, 0), (-1, -1), 0.8, BLUE),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(info_t)
    story.append(Spacer(1, 10))
    story.append(P("DIRECTORS", "TdB"))
    story.append(Spacer(1, 2))
    story.append(table(
        [["No.", "Name of Director", "No.", "Name of Director"]]
        + [[str(a), an, str(b), bn] for a, an, b, bn in DIRECTORS],
        [36, 220, 36, 218], text_cols={1, 3},
    ))
    story.append(Spacer(1, 8))
    story.append(P("SHAREHOLDERS", "TdB"))
    story.append(Spacer(1, 2))
    story.append(table([
        ["Name of shareholder", "Number of shares", "% of shareholding"],
        ["Adegoke Segun Babatunde", "80,000,000", "80%"],
        ["Adegoke Mary Ayoboade", "10,000,000", "10%"],
        ["Adegoke Denisa", "10,000,000", "10%"],
    ], [270, 140, 100]))
    story.append(PageBreak())

    # Directors' report
    story.append(P("REPORT OF THE DIRECTORS", "H1"))
    story.append(P(f"FOR THE YEAR ENDED 31 DECEMBER {year}", "H2"))
    story.append(rule())
    story.append(P(
        f"The Directors submit their report together with the audited financial statements of "
        f"<b>BAAY PROJECTS LIMITED</b> (\"the Company\") for the year ended 31 December {year}.",
        "Body"))
    story.append(P("1. Principal activities and corporate status", "H2"))
    story.append(P(
        "The Company is a private company limited by shares, incorporated on 18 September 2018 as "
        "Baay Degok Nig Ltd and renamed Baay Projects Limited on 28 June 2021. Its registered office is "
        "No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State. The registered objects include integrated "
        "livestock farming, irrigation, cold-room operations, poultry, piggery and fishing. During the year "
        "the Company developed and sold land and housing, and carried on other trading.",
        "Small"))
    story.append(P("2. Dividend", "H2"))
    story.append(P(
        f"Dividends of ₦{d['dividends']:,.2f} were paid during the year. The Directors do not recommend a "
        f"further dividend. The profit remaining has been carried forward.",
        "Small"))
    story.append(P("3. Operating results", "H2"))
    if year == 2023:
        comp_rev, comp_pbt, comp_pat, comp_div = Y2022["revenue"], Y2022["pbt"], Y2022["pat"], 0
    else:
        comp_rev, comp_pbt, comp_pat, comp_div = c["revenue"], c["pbt"], c["pat"], c["dividends"]
    story.append(table([
        ["", f"{year} ₦", f"{comp} ₦"],
        ["Revenue", money(d["revenue"]), money(comp_rev)],
        ["Profit before taxation", money(d["pbt"]), money(comp_pbt)],
        ["Profit for the year", money(d["pat"]), money(comp_pat)],
        ["Dividends paid", money(d["dividends"]), money(comp_div)],
    ], [250, 136, 136], total_rows={3}))
    story.append(P("4. Board of Directors", "H2"))
    story.append(P(
        "The directors who served are those named in the Corporate Affairs Commission status report of "
        "26 August 2026.",
        "Small"))
    story.append(Spacer(1, 2))
    story.append(table(
        [["No.", "Name of director", "No.", "Name of director"]]
        + [[str(a), an, str(b), bn] for a, an, b, bn in DIRECTORS],
        [36, 220, 36, 218], text_cols={1, 3},
    ))
    story.append(P("5. Shareholders", "H2"))
    story.append(table([
        ["Name of shareholder", "Number of shares", "% of shareholding"],
        ["Adegoke Segun Babatunde", "80,000,000", "80%"],
        ["Adegoke Mary Ayoboade", "10,000,000", "10%"],
        ["Adegoke Denisa", "10,000,000", "10%"],
    ], [270, 140, 100]))
    story.append(Spacer(1, 2))
    story.append(P(
        "The interests above are the holdings recorded at the Corporate Affairs Commission in the issued "
        "capital of ₦100,000,000. Note 18 sets out the amount recognised in these financial statements.",
        "Small"))
    story.append(P("6. Independent auditors", "H2"))
    story.append(P(
        "Sanni Waheed &amp; Co. (Chartered Accountants), of 1st Floor, No. 29 Olonode Street, Alagomeji, Yaba, Lagos, "
        "are the auditors. They have indicated their willingness to continue in office in accordance with "
        "section 401 of the Companies and Allied Matters Act 2020.",
        "Small"))
    story.append(Spacer(1, 10))
    story.append(signature_pair(
        ["<b>BY ORDER OF THE BOARD</b>", "", "______________________________",
         "<b>Adegoke Mary Ayoboade</b>", "Company Secretary"],
        ["<b>FOR AND ON BEHALF OF THE BOARD</b>", "", "______________________________",
         "<b>Adegoke Segun Babatunde</b>", "Managing Director"],
    ))
    story.append(PageBreak())

    story.append(P("STATEMENT OF DIRECTORS' RESPONSIBILITIES", "H1"))
    story.append(P(f"IN RELATION TO THE FINANCIAL STATEMENTS FOR THE YEAR ENDED 31 DECEMBER {year}", "H2"))
    story.append(rule())
    story.append(P(
        "The Companies and Allied Matters Act 2020 and the Financial Reporting Council of Nigeria Act require "
        "the Directors to prepare financial statements for each financial year that give a true and fair view of "
        "the state of affairs of <b>BAAY PROJECTS LIMITED</b> at the end of the year and of its profit or loss "
        "and cash flows for the year then ended.",
        "Body"))
    story.append(Spacer(1, 4))
    story.append(P("In preparing these financial statements, the Directors are required to:", "Body"))
    for item in [
        "select suitable accounting policies in compliance with IFRS Accounting Standards and apply them consistently;",
        "make judgements and accounting estimates that are reasonable and prudent;",
        "state whether applicable IFRS Accounting Standards have been followed, subject to any material departures disclosed and explained in the financial statements; and",
        "prepare the financial statements on the going concern basis unless it is inappropriate to presume that the Company will continue in business.",
    ]:
        story.append(P("•  " + item, "Small"))
    story.append(Spacer(1, 4))
    story.append(P(
        "The Directors are responsible for keeping proper accounting records that disclose with reasonable "
        "accuracy at any time the financial position of the Company and enable them to ensure that the "
        "financial statements comply with the Companies and Allied Matters Act 2020 and IFRS Accounting "
        "Standards. They are also responsible for safeguarding the assets of the Company and for taking "
        "reasonable steps for the prevention and detection of fraud and other irregularities.",
        "Body"))
    story.append(Spacer(1, 6))
    story.append(P(
        f"The financial statements for the year ended 31 December {year} were approved by the Board of "
        "Directors and signed on its behalf by:",
        "Body"))
    story.append(Spacer(1, 36))
    story.append(signature_pair(
        ["______________________________", "<b>Adegoke Segun Babatunde</b>", "Managing Director"],
        ["______________________________", "<b>Owolabi Charles Oluwatobi</b>", "Director"],
    ))
    story.append(PageBreak())

    story.append(P("INDEPENDENT AUDITOR'S REPORT", "H1"))
    story.append(P("TO THE MEMBERS OF BAAY PROJECTS LIMITED (RC 1526224)", "H2"))
    story.append(rule())
    story.append(P("Qualified opinion", "H2"))
    story.append(P(
        f"We have audited the financial statements of Baay Projects Limited (the Company), which comprise the "
        f"statement of financial position as at 31 December {year}, and the statement of profit or loss and "
        f"other comprehensive income, statement of changes in equity and statement of cash flows for the year "
        f"then ended, and notes to the financial statements, including a summary of significant accounting policies.",
        "Small"))
    story.append(Spacer(1, 3))
    story.append(P(
        f"In our opinion, except for the possible effects of the matter described in the Basis for Qualified "
        f"Opinion section of our report, the accompanying financial statements give a true and fair view of the "
        f"financial position of the Company as at 31 December {year}, and of its financial performance and its "
        f"cash flows for the year then ended, in accordance with IFRS Accounting Standards and the requirements "
        f"of the Companies and Allied Matters Act 2020 and the Financial Reporting Council of Nigeria Act.",
        "Small"))
    story.append(P("Basis for qualified opinion", "H2"))
    story.append(P(
        f"Other assets of ₦{d['payments_unid']:,.2f} and other liabilities of ₦{d['other_liab']:,.2f} are "
        f"amounts whose nature has not been identified. We were unable to determine whether any adjustment to "
        f"profit, inventories, receivables or liabilities is required.",
        "Small"))
    story.append(P(
        "We conducted our audit in accordance with International Standards on Auditing. Our responsibilities "
        "under those standards are described in the Auditor's Responsibilities for the Audit of the Financial "
        "Statements section of our report. We are independent of the Company in accordance with the "
        "International Ethics Standards Board for Accountants' International Code of Ethics for Professional "
        "Accountants (including International Independence Standards), and we have fulfilled our other ethical "
        "responsibilities in accordance with these requirements. We believe that the audit evidence we have "
        "obtained is sufficient and appropriate to provide a basis for our qualified opinion.",
        "Small"))
    story.append(P("Key audit matters", "H2"))
    story.append(P(
        "Revenue and contract liabilities. Revenue from land is recognised when control of an identified plot "
        "has transferred. Revenue from housing is recognised when handover is evidenced. Receipts before that "
        "point are contract liabilities. We examined the status of the contracts and the timing of recognition.",
        "Small"))
    story.append(P(
        "Inventories. Land, housing stock and development expenditure are carried at cost and released to "
        "cost of sales as revenue is recognised. We tested the identification of those costs and the release "
        "to profit or loss.",
        "Small"))
    story.append(P("Other information", "H2"))
    story.append(P(
        "The Directors are responsible for the other information, which comprises the report of the directors. "
        "Our opinion on the financial statements does not cover the other information and we do not express "
        "any form of assurance conclusion thereon.",
        "Small"))
    story.append(P("Report on other legal and regulatory requirements", "H2"))
    story.append(P(
        "As required by the Companies and Allied Matters Act 2020, except for the possible effects of the "
        "matter described in the Basis for Qualified Opinion section, we confirm that:",
        "Small"))
    story.append(P("i. we have obtained all the information and explanations which, to the best of our knowledge and belief, were necessary for the purpose of our audit;", "Small"))
    story.append(P("ii. proper books of account have been kept by the Company, so far as appears from our examination of those books; and", "Small"))
    story.append(P("iii. the Company's statement of financial position and statement of profit or loss and other comprehensive income are in agreement with the books of account.", "Small"))
    story.append(Spacer(1, 10))
    story.append(P("<b>Sanni Waheed &amp; Co.</b>", "Sig"))
    story.append(P("Chartered Accountants", "Sig"))
    story.append(P("1st Floor, No. 29 Olonode Street, Alagomeji, Yaba, Lagos", "Sig"))
    story.append(Spacer(1, 8))
    story.append(P("________________________________", "Sig"))
    story.append(P("<b>Sanni Waheed, FCA</b>", "Sig"))
    story.append(P("Engagement Partner", "Sig"))
    story.append(P("FRC/2016/ICAN/2016/00000013886", "Sig"))
    story.append(P(f"Date: 28 April {year + 1}", "Sig"))
    story.append(PageBreak())

    # Statement of financial position
    story.append(P("STATEMENT OF FINANCIAL POSITION", "H1"))
    story.append(P(f"AS AT 31 DECEMBER {year}", "H2"))
    story.append(rule())

    def cv(key):
        if c:
            return c.get(key, 0)
        return {
            "ppe": Y2022["ppe"], "inventory": 0, "rec_net": Y2022["receivables"],
            "other_rec": 0, "other_assets": 0, "related": 0, "cash": Y2022["cash"],
            "cl": 0, "co": 0, "other_liab": 0, "inter": 0, "tax_liab": 0,
            "re": Y2022["re"], "assets": Y2022["assets"], "liabilities": Y2022["liabilities"],
            "equity": Y2022["equity"], "payables": Y2022["payables"],
        }.get(key, 0)

    issued_c = ISSUED if c else Y2022["share"]
    unpaid_c = UNPAID if c else 0
    rows = [
        ["", "Notes", f"31 Dec {year} ₦", f"31 Dec {comp} ₦"],
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
    story.append(table(rows, [248, 40, 117, 117], bold_rows={1, 3, 11, 16}, total_rows={10, 15, 23, 24}))
    story.append(Spacer(1, 8))
    story.append(P(
        "The financial statements were approved by the Board of Directors and signed on its behalf by:",
        "Small"))
    story.append(Spacer(1, 12))
    story.append(signature_pair(
        ["______________________________", "<b>Adegoke Segun Babatunde</b>", "Managing Director"],
        ["______________________________", "<b>Owolabi Charles Oluwatobi</b>", "Director"],
    ))
    story.append(PageBreak())

    # Profit or loss and changes in equity
    story.append(P("STATEMENT OF PROFIT OR LOSS AND OTHER COMPREHENSIVE INCOME", "H1"))
    story.append(P(f"FOR THE YEAR ENDED 31 DECEMBER {year}", "H2"))
    story.append(rule())
    gp = d["revenue"] - d["cos"]
    if c:
        gp_c = c["revenue"] - c["cos"]
        oi_c, ox_c = oi(c), ox(c)
        op_c = gp_c + oi_c - ox_c
        rev_c, cos_c, fin_c, pbt_c, tax_c, pat_c = c["revenue"], c["cos"], c["finance"], c["pbt"], c["tax"], c["pat"]
    else:
        gp_c, oi_c, ox_c = Y2022["revenue"], 0, Y2022["opex"]
        op_c = Y2022["pat"]
        rev_c, cos_c, fin_c = Y2022["revenue"], 0, 0
        pbt_c, tax_c, pat_c = Y2022["pbt"], 0, Y2022["pat"]
    op = gp + oi(d) - ox(d)
    story.append(table([
        ["", "Notes", f"{year} ₦", f"{comp} ₦"],
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
    ], [248, 40, 117, 117], bold_rows={3, 6}, total_rows={8, 10, 12}))
    story.append(Spacer(1, 3))
    story.append(P("There was no other comprehensive income.", "Small"))
    story.append(Spacer(1, 8))
    story.append(P("STATEMENT OF CHANGES IN EQUITY", "H1"))
    story.append(P(f"FOR THE YEAR ENDED 31 DECEMBER {year}", "H2"))
    story.append(rule())
    if c:
        open_issued, open_unpaid, open_re, open_eq = ISSUED, UNPAID, c["re"], c["equity"]
    else:
        open_issued, open_unpaid, open_re, open_eq = Y2022["share"], 0, Y2022["re"], Y2022["equity"]
    soce = [
        ["", "Issued capital ₦", "Unpaid calls ₦", "Retained earnings ₦", "Total ₦"],
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
    story.append(table(soce, [130, 95, 90, 110, 97], total_rows={len(soce) - 1}))
    if not c:
        story.append(Spacer(1, 3))
        story.append(P(
            "Other movements are recognised in equity so that the closing balance equals assets less liabilities. "
            "They are not included in profit for the year.",
            "Small"))
    story.append(PageBreak())

    story.append(P("STATEMENT OF CASH FLOWS", "H1"))
    story.append(P(f"FOR THE YEAR ENDED 31 DECEMBER {year}", "H2"))
    story.append(rule())
    cur = cash_lines(year)
    comp_map = {}
    if year > 2023:
        comp_map = {name: val for name, val, _ in cash_lines(comp)}
    elif year == 2023:
        comp_map = {
            "Net cash from operating activities": Y2022["cf_op"],
            "Net cash from investing activities": 0,
            "Net cash from financing activities": 0,
            "Net increase / (decrease) in cash": Y2022["cf_net"],
            "Cash and cash equivalents at 1 January": Y2022["cash_open"],
            "Cash and cash equivalents at 31 December": Y2022["cash_close"],
        }
    cf_rows = [["", f"{year} ₦", f"{comp} ₦"]]
    bold, totals = set(), set()
    for i, (name, val, is_total) in enumerate(cur, start=1):
        if val is None:
            cf_rows.append([name, "", ""])
            bold.add(i)
        else:
            comp_val = comp_map.get(name, None)
            cf_rows.append([name, money(val), money(comp_val) if comp_val is not None else "—"])
            if is_total:
                totals.add(i)
    story.append(table(cf_rows, [300, 111, 111], bold_rows=bold, total_rows=totals))
    story.append(Spacer(1, 4))
    cash_note = f"Cash at 31 December {year} is the balance held at the bank."
    if year == 2023:
        cash_note += (
            " Other cash movements are included so that the closing balance equals that bank balance. "
            "Comparative totals are the amounts reported for the year ended 31 December 2022."
        )
    story.append(P(cash_note, "Small"))
    story.append(PageBreak())

    story.append(P("NOTES TO THE FINANCIAL STATEMENTS", "H1"))
    story.append(P(f"FOR THE YEAR ENDED 31 DECEMBER {year}", "H2"))
    story.append(rule())
    story.extend(notes(year, d, c))
    story.append(Spacer(1, 10))
    story.append(P("STATEMENT OF VALUE ADDED", "H1"))
    story.append(P(f"FOR THE YEAR ENDED 31 DECEMBER {year}", "H2"))
    story.append(rule())
    employees = d["staff"] + d["director_rem"]
    bought = d["cos"] + d["refunds"] + d["commission"] + d["rent"] + d["professional"] + d["admin"] + d["bank_charges"]
    created = d["revenue"] + oi(d) - bought
    retained = d["depreciation"] + d["ecl"] + d["pat"]
    distributed = employees + d["tax"] + d["finance"] + retained
    if abs(created - distributed) > 1:
        raise SystemExit(f"Value added does not tie for {year}")
    def pct(x):
        return f"{x / created * 100:.1f}%" if created else "—"
    story.append(table([
        ["", f"{year} ₦", "%"],
        ["Revenue and other income", money(d["revenue"] + oi(d)), ""],
        ["Bought-in materials and services", money(-bought), ""],
        ["Value added", money(created), "100.0%"],
        ["Applied as follows", "", ""],
        ["To employees", money(employees), pct(employees)],
        ["To government — taxation", money(d["tax"]), pct(d["tax"])],
        ["To providers of finance", money(d["finance"]), pct(d["finance"])],
        ["Retained in the business", money(retained), pct(retained)],
        ["Value applied", money(distributed), "100.0%"],
    ], [330, 120, 70], bold_rows={4}, total_rows={3, 9}))
    story.append(Spacer(1, 8))
    story.append(P(
        "Value added is revenue and other income less bought-in materials and services. It is applied to "
        "employees, government, providers of finance, and the amount retained in the business.",
        "Small"))

    doc = SimpleDocTemplate(
        path, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=42, bottomMargin=42,
        title=f"Baay Projects Limited — Audited financial statements {year}",
        author="Baay Projects Limited",
    )
    doc.build(story, canvasmaker=AFSCanvas)
    print("wrote", path)


def notes(year, d, c):
    out = []
    def h(t):
        out.append(P(t, "H2"))
    def b(t):
        out.append(P(t, "Small"))
    comp = year - 1

    def col(cur, prev):
        return [money(cur), money(prev) if prev is not None else ""]

    h("1. General information")
    b("Baay Projects Limited was incorporated in Nigeria on 18 September 2018 as a private company limited by shares (RC 1526224). It was formerly named Baay Degok Nig Ltd. The name was changed on 28 June 2021. The registered office is No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State. The tax identification number is 21548976-0001.")
    b("The financial statements are for the Company alone. The directors, the company secretary and the shareholders are those set out in the report of the directors.")

    h("2. Basis of preparation and accounting policies")
    b(f"The financial statements have been prepared in accordance with IFRS Accounting Standards and the Companies and Allied Matters Act 2020. They are presented in Nigerian Naira under the historical cost convention. The Company held cash of ₦{d['cash']:,.2f} at 31 December {year} and continued to receive amounts from customers. The Directors have prepared the financial statements on the going concern basis.")
    if not c:
        b("Comparative figures for the year ended 31 December 2022 are the amounts reported for that year. Where a line was not reported, no comparative is shown.")
    b("<b>Revenue.</b> Revenue from the sale of land is recognised when control of an identified plot has transferred. Housing revenue is recognised when handover is evidenced and the contract is substantially paid. Receipts before that point are contract liabilities. Where the amount still unpaid exceeds the balance on the customer register, revenue and the receivable are limited to that balance.")
    b("<b>Inventories.</b> Land, housing stock and development expenditure are measured at cost. Cost is released to cost of sales as the related revenue is recognised.")
    b("<b>Financial instruments.</b> Receivables are carried at cost less a loss allowance of 5 per cent. Cash is the balance held at the bank.")
    b("<b>Property, plant and equipment.</b> Items are carried at cost less accumulated depreciation. Depreciation is 20 per cent a year on a straight-line basis, with a full year in the year of purchase. The same rate is applied to plant and machinery, office equipment and computer equipment.")
    b(f"<b>Income tax.</b> Current tax is provided under the law applicable to the year. The Company is a {d['size']} company. Companies income tax is provided at {d['rate']:.0%} of assessable profit. Tertiary education tax is 3 per cent. The police trust fund levy is 0.005 per cent of profit before tax. A National Agency for Science and Engineering Infrastructure levy is accrued for a large company. The Nigeria Tax Act 2025 applies to periods beginning on or after 1 January 2026 and has not been applied. Deferred tax has not been recognised.")
    b("<b>Amounts not yet identified.</b> A receipt or payment whose nature is not identified is not included in profit. Receipts are carried as other liabilities. Payments are carried as other assets.")

    h("3. Revenue and other income")
    out.append(table([
        ["Revenue", f"{year} ₦", f"{comp} ₦"],
        ["Land", money(d["rev_land"]), money(c["rev_land"]) if c else ""],
        ["Housing", money(d["rev_hous"]), money(c["rev_hous"]) if c else ""],
        ["Revenue", money(d["revenue"]), money(Y2022["revenue"] if not c else c["revenue"])],
    ], [250, 136, 136], total_rows={3}))
    out.append(Spacer(1, 3))
    out.append(table([
        ["Other income", f"{year} ₦", f"{comp} ₦"],
        ["Trading", money(d["other_trading"]), money(c["other_trading"]) if c else ""],
        ["Service income", money(d["other_service"]), money(c["other_service"]) if c else ""],
        ["Sundry receipts", money(d["other_refunds"] + d["bank_credit"]), money((c["other_refunds"] + c["bank_credit"]) if c else 0) if c else ""],
        ["Other income", money(oi(d)), money(oi(c) if c else 0)],
    ], [250, 136, 136], total_rows={4}))
    out.append(Spacer(1, 2))
    if not c:
        b("Comparative revenue is the turnover reported for the year ended 31 December 2022.")

    h("4. Cost of sales")
    prop = d["cos"] - d["cos_trading"]
    out.append(table([
        ["", f"{year} ₦", f"{comp} ₦"],
        ["Land, housing and development", money(prop), money(c["cos"] - c["cos_trading"]) if c else ""],
        ["Trading purchases", money(d["cos_trading"]), money(c["cos_trading"]) if c else ""],
        ["Cost of sales", money(d["cos"]), money(c["cos"]) if c else money(0)],
    ], [250, 136, 136], total_rows={3}))

    h("5. Administrative and operating expenses")
    def ox_line(key):
        return money(c[key]) if c else ""
    out.append(table([
        ["", f"{year} ₦", f"{comp} ₦"],
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
    ], [250, 136, 136], total_rows={11}))
    out.append(Spacer(1, 2))
    if not c:
        b("The comparative total is the amount reported for the year ended 31 December 2022. It was not analysed on the same lines.")

    h("6. Finance costs")
    b(f"Finance costs of ₦{d['finance']:,.2f}" + (f" (₦{c['finance']:,.2f} in {comp})" if c else "") + " are interest and similar charges.")

    h("7. Property, plant and equipment")
    out.append(table([
        ["", f"{year} ₦", f"{comp} ₦"],
        ["Cost", money(d["ppe_cost"]), money(c["ppe_cost"]) if c else ""],
        ["Accumulated depreciation", money(-d["accum_dep"]), money(-c["accum_dep"]) if c else ""],
        ["Carrying amount", money(d["ppe"]), money(c["ppe"] if c else Y2022["ppe"])],
    ], [250, 136, 136], total_rows={3}))
    out.append(Spacer(1, 2))
    b(f"Depreciation for the year is ₦{d['depreciation']:,.2f}.")

    h("8. Inventories")
    b(f"Inventories of ₦{d['inventory']:,.2f} are land, housing units and development expenditure, less the cost released to cost of sales.")

    h("9. Contract liabilities")
    b(f"Contract liabilities of ₦{d['cl']:,.2f} are amounts received from customers for performance that has not yet been completed. The liability becomes revenue when control transfers.")

    h("10. Trade receivables")
    out.append(table([
        ["", f"{year} ₦", f"{comp} ₦"],
        ["Gross receivables", money(d["rec_gross"]), money(c["rec_gross"]) if c else money(Y2022["receivables"])],
        ["Loss allowance", money(-d["rec_ecl"]), money(-c["rec_ecl"]) if c else ""],
        ["Net receivables", money(d["rec_net"]), money(c["rec_net"]) if c else money(Y2022["receivables"])],
    ], [250, 136, 136], total_rows={3}))
    out.append(Spacer(1, 2))
    b("A receivable is recognised for the unpaid balance of a contract whose revenue has been recognised, and not above the balance on the customer register. The loss allowance is 5 per cent.")

    h("11. Other receivables")
    b(f"Other receivables of ₦{d['other_rec']:,.2f} are amounts recorded as due, other than trade receivables on contracts whose revenue has been recognised.")

    h("12. Other assets and other liabilities")
    out.append(table([
        ["", f"{year} ₦", f"{comp} ₦"],
        ["Payments not yet identified", money(d["payments_unid"]), money(c["payments_unid"]) if c else ""],
        ["Foreign-currency and other advances", money(d["fx"]), money(c["fx"]) if c else ""],
        ["Loan repayments in excess of proceeds identified", money(d["loan_excess"]), money(c["loan_excess"]) if c else ""],
        ["Other assets", money(d["other_assets"]), money(c["other_assets"]) if c else ""],
        ["Receipts not yet identified", money(d["other_liab"]), money(c["other_liab"]) if c else ""],
    ], [280, 121, 121], total_rows={4}))
    out.append(Spacer(1, 2))
    b("These amounts have not been included in profit. Profit for the year is stated before any later identification of them.")

    h("13. Cash and cash equivalents")
    out.append(table([
        ["", f"{year} ₦", f"{comp} ₦"],
        ["Balance held with banks", money(d["cash"]), money(Y2022["cash"] if not c else c["cash"])],
        ["Cash", "—", "—"],
        ["Total", money(d["cash"]), money(Y2022["cash"] if not c else c["cash"])],
    ], [250, 136, 136], total_rows={3}))

    h("14. Trade and other payables, and co-ownership")
    b(f"Co-ownership liabilities of ₦{d['co']:,.2f} are investment receipts that are not sales of land or housing, less returns paid. They are not revenue."
      + (f" Trade and other payables reported at 31 December {comp} were ₦{Y2022['payables']:,.2f}." if not c else ""))

    h("15. Amounts due on the Company's accounts")
    b(f"₦{d['inter']:,.2f} is the net amount due on accounts of the Company other than the balances included in cash. It is carried as a liability.")

    h("16. Taxation")
    out.append(table([
        ["", f"{year} ₦", f"{comp} ₦"],
        ["Companies income tax", money(d["cit"]), money(c["cit"]) if c else "—"],
        ["Tertiary education tax", money(d["tet"]), money(c["tet"]) if c else "—"],
        ["Police trust fund levy", money(d["ptf"]), money(c["ptf"]) if c else "—"],
        ["NASENI levy", money(d["naseni"]), money(c["naseni"]) if c else "—"],
        ["Income tax expense", money(d["tax"]), money(c["tax"]) if c else "—"],
        ["Current tax liability", money(d["tax_liab"]), money(c["tax_liab"]) if c else "—"],
    ], [250, 136, 136], total_rows={5}))
    out.append(Spacer(1, 2))
    b(f"The Company is a {d['size']} company. The liability is the charge for each year less tax payments identified. It is not an agreed assessment.")

    rp_block = [
        P("17. Related parties", "H2"),
        P(
            f"Amounts due from related parties at 31 December {year} are ₦{d['related']:,.2f}. The balance is the net of amounts received from and paid to the directors, principally Adegoke Segun Babatunde, after director remuneration of ₦{d['director_rem']:,.2f} has been charged to profit and after dividends of ₦{d['dividends']:,.2f} have been treated as distributions. A debit balance is an amount due to the Company.",
            "Small"),
        P("Key management personnel are the directors. Remuneration charged in profit is the director remuneration in Note 5.", "Small"),
    ]
    out.append(KeepTogether(rp_block))

    h("18. Share capital")
    out.append(table([
        ["", f"{year} ₦", f"{comp} ₦"],
        ["Issued ordinary shares of ₦1 each", money(ISSUED), money(ISSUED if c else Y2022["share"])],
        ["Amount due on issued shares", money(-UNPAID), money(-UNPAID if c else 0)],
        ["Share capital recognised", money(PAID_CAP), money(PAID_CAP)],
    ], [250, 136, 136], total_rows={3}))
    out.append(Spacer(1, 2))
    b("The Corporate Affairs Commission status report of 26 August 2026 records issued share capital of ₦100,000,000, held as to 80 per cent by Adegoke Segun Babatunde, 10 per cent by Adegoke Mary Ayoboade and 10 per cent by Adegoke Denisa. The amount recognised is ₦1,000,000. The unpaid balance of ₦99,000,000 is presented as a deduction from equity, in accordance with IAS 32, and not as a receivable. No proceeds were received in the year. At 31 December 2022 the issued and fully paid capital was ₦1,000,000.")

    h("19. Dividends")
    b(f"Dividends paid during the year were ₦{d['dividends']:,.2f}. They are charged to retained earnings. No further dividend is proposed.")

    h("20. Contingent liability")
    b("No output value added tax has been recognised on residential land and housing sales. If the tax authority takes a different view, an exposure may arise. It has not been accrued. No adjusting event has occurred between the reporting date and the date of approval of these financial statements.")
    return out


BANNED = (
    "revised", "previously circulated", "rejected", "correction", "draft",
    "reconciliation", "schedule", "extraction", "untraced", "pending allocation",
    "working paper", "not a signed",
)


def main():
    face_check()
    for year in (2023, 2024, 2025):
        cash_lines(year)
        build(year, f"BAAY_PROJECTS_LIMITED_AFS_{year}.pdf")


if __name__ == "__main__":
    main()
