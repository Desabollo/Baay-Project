"""Write the four revised workbooks from the engine model."""

from __future__ import annotations

import pickle
import shutil
from pathlib import Path

import pandas as pd
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path("/home/user/Baay-Project")
MODEL = Path("/tmp/baay_extract/model.pkl")

OLD = {
    2023: dict(revenue=38_500_000, cash=2_842_403, receivables=5_015_000, deferred=2_400_000, assets=20_363_943, pbt=7_300_000),
    2024: dict(revenue=64_200_000, cash=5_142_678, receivables=8_685_000, deferred=3_600_000, assets=35_187_218, pbt=12_500_000),
    2025: dict(revenue=112_500_000, cash=8_924_048, receivables=14_360_000, deferred=5_800_000, assets=60_158_588, pbt=22_600_000),
}


def xf(wb):
    n = {"num_format": "#,##0.00", "border": 1}
    return {
        "title": wb.add_format({"bold": True, "font_size": 16, "font_color": "1F4E79", "font_name": "Calibri"}),
        "sub": wb.add_format({"font_size": 11, "font_color": "333333", "text_wrap": True, "font_name": "Calibri"}),
        "h": wb.add_format({"bold": True, "bg_color": "1F4E79", "font_color": "white", "border": 1, "text_wrap": True, "valign": "vcenter", "font_name": "Calibri"}),
        "sec": wb.add_format({"bold": True, "font_size": 12, "font_color": "1F4E79", "font_name": "Calibri"}),
        "txt": wb.add_format({"border": 1, "valign": "vcenter", "text_wrap": True, "font_name": "Calibri", "font_size": 10}),
        "txtb": wb.add_format({"border": 1, "bold": True, "bg_color": "E7EEF6", "font_name": "Calibri", "font_size": 10}),
        "m": wb.add_format({**n, "font_name": "Calibri", "font_size": 10}),
        "mb": wb.add_format({**n, "bold": True, "bg_color": "E7EEF6", "font_name": "Calibri"}),
        "i": wb.add_format({"border": 1, "num_format": "#,##0", "font_name": "Calibri", "font_size": 10}),
        "note": wb.add_format({"text_wrap": True, "font_size": 10, "valign": "top", "font_name": "Calibri"}),
        "warn": wb.add_format({"text_wrap": True, "font_size": 10, "font_color": "8C2F2F", "bold": True, "font_name": "Calibri"}),
        "ok": wb.add_format({"bold": True, "font_color": "1E6B4F", "font_name": "Calibri", "border": 1}),
        "bad": wb.add_format({"bold": True, "font_color": "8C2F2F", "font_name": "Calibri", "border": 1}),
        "date": wb.add_format({"num_format": "yyyy-mm-dd", "border": 1, "font_name": "Calibri", "font_size": 9}),
    }


def cover(ws, F, title, lines):
    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 110)
    ws.write(0, 0, "BAAY PROJECTS LIMITED  ·  RC 1526224", F["sec"])
    ws.write(1, 0, title, F["title"])
    ws.write(2, 0, "REVISED  ·  8 October 2026  ·  amounts in Nigerian Naira", F["warn"])
    for i, line in enumerate(lines):
        ws.write(4 + i, 0, line, F["note"])
        ws.set_row(4 + i, 32)
    ws.set_row(1, 24)


def write_bank(model):
    path = ROOT / "Baay_Projects_Clean_Bank_Statements_Extraction_2023_2025_revised.xlsx"
    wb = xlsxwriter.Workbook(str(path), {"nan_inf_to_errors": True})
    F = xf(wb)
    b = model["bank"].copy()
    b["date"] = pd.to_datetime(b["date"])

    ws = wb.add_worksheet("Cover")
    cover(ws, F, "Clean bank statements extraction — revised", [
        "This file replaces the extraction as the working cash book for the revision. Every transaction from 1 January 2023 to 31 December 2025 on the four accounts is classified.",
        "Accounts: Providus 5400724970, Sterling 0091166190, First Bank 2033736938 (still in the former name Baay Degok Nig Ltd), GTBank 0893079582. GTBank opened for activity in 2024.",
        "The account numbers in the previously circulated financial statements are not these accounts and are not used.",
        "Cash on the revised statements is the stated balance on these four accounts. The cash proof on the next sheet must be nil.",
        "A narration that does not support a classification is left in Receipts pending allocation or Payments pending allocation. It is not forced into revenue or into profit.",
        "Inter-account transfers between these four accounts are classified so they do not create income. Transfers that do not mirror inside the four accounts remain visible as a financing balance.",
    ])

    ws = wb.add_worksheet("Cash_proof")
    ws.set_column(0, 0, 22)
    ws.set_column(1, 6, 20)
    headers = ["Bank", "Opening 1 Jan 2023", "2023 stated", "2024 stated", "2025 stated", "Name on statement", "Account"]
    for c, h in enumerate(headers):
        ws.write(0, c, h, F["h"])
    meta = {
        "GTBank": ("BAAY PROJECTS LTD", "0893079582"),
        "First Bank": ("BAAY DEGOK NIG LTD", "2033736938"),
        "Sterling": ("BAAY PROJECTS", "0091166190"),
        "Providus": ("BAAY PROJECTS LTD", "5400724970"),
    }
    for r, bk in enumerate(["Providus", "Sterling", "First Bank", "GTBank"], start=1):
        ws.write(r, 0, bk, F["txt"])
        ws.write_number(r, 1, model["stated_cash"][2022]["by_bank"][bk], F["m"])
        for i, y in enumerate((2023, 2024, 2025)):
            ws.write_number(r, 2 + i, model["stated_cash"][y]["by_bank"][bk], F["m"])
        ws.write(r, 5, meta[bk][0], F["txt"])
        ws.write(r, 6, meta[bk][1], F["txt"])
    ws.write(5, 0, "Total", F["txtb"])
    for c, y in enumerate((2022, 2023, 2024, 2025)):
        ws.write_number(5, 1 + c, model["stated_cash"][y]["total"], F["mb"])
    ws.write(7, 0, "Movement agrees to classified transactions", F["sec"])
    ws.write(8, 0, "Year")
    ws.write(8, 1, "Opening", F["h"])
    ws.write(8, 2, "Credits", F["h"])
    ws.write(8, 3, "Debits", F["h"])
    ws.write(8, 4, "Computed close", F["h"])
    ws.write(8, 5, "Stated close", F["h"])
    ws.write(8, 6, "Difference", F["h"])
    ws.write(8, 0, "Year", F["h"])
    open_cash = model["opening_cash_total"]
    for i, y in enumerate((2023, 2024, 2025)):
        sub = b[b["year"] == y]
        cr, dr = float(sub["cr"].sum()), float(sub["dr"].sum())
        close = open_cash + cr - dr
        stated = model["stated_cash"][y]["total"]
        ws.write(9 + i, 0, y, F["txt"])
        ws.write_number(9 + i, 1, open_cash, F["m"])
        ws.write_number(9 + i, 2, cr, F["m"])
        ws.write_number(9 + i, 3, dr, F["m"])
        ws.write_number(9 + i, 4, close, F["m"])
        ws.write_number(9 + i, 5, stated, F["m"])
        diff = close - stated
        ws.write_number(9 + i, 6, diff, F["mb"])
        ws.write(9 + i, 7, "PASS" if abs(diff) < 1 else "FAIL", F["ok"] if abs(diff) < 1 else F["bad"])
        open_cash = stated
    ws.write(13, 0, "A difference other than nil means this file must not be used. The engine requires nil.", F["note"])

    ws = wb.add_worksheet("Classification")
    ws.set_column(0, 0, 36)
    ws.set_column(1, 8, 18)
    ws.write(0, 0, "Category", F["h"])
    for i, y in enumerate((2023, 2024, 2025)):
        ws.write(0, 1 + i * 2, f"{y} credits", F["h"])
        ws.write(0, 2 + i * 2, f"{y} debits", F["h"])
    ws.write(0, 7, "Credits 2023-25", F["h"])
    ws.write(0, 8, "Debits 2023-25", F["h"])
    cats = sorted(b["category"].unique())
    for r, cat in enumerate(cats, start=1):
        ws.write(r, 0, cat, F["txt"])
        for i, y in enumerate((2023, 2024, 2025)):
            sub = b[(b["year"] == y) & (b["category"] == cat)]
            ws.write_number(r, 1 + i * 2, float(sub["cr"].sum()), F["m"])
            ws.write_number(r, 2 + i * 2, float(sub["dr"].sum()), F["m"])
        sub = b[b["category"] == cat]
        ws.write_number(r, 7, float(sub["cr"].sum()), F["m"])
        ws.write_number(r, 8, float(sub["dr"].sum()), F["m"])
    ws.autofilter(0, 0, len(cats), 8)
    ws.freeze_panes(1, 1)

    ws = wb.add_worksheet("Transactions")
    cols = ["txn_id", "bank", "acct", "acct_name", "date", "ref", "narr",
            "dr", "cr", "bal", "year", "category", "bucket"]
    for c, h in enumerate(cols):
        ws.write(0, c, h, F["h"])
    widths = [14, 14, 16, 28, 12, 22, 70, 16, 16, 16, 8, 28, 22]
    for c, w in enumerate(widths):
        ws.set_column(c, c, w)
    b = b.sort_values(["date", "bank", "txn_id"])
    recs = b[cols]
    for r, rec in enumerate(recs.itertuples(index=False), start=1):
        ws.write(r, 0, rec.txn_id, F["txt"])
        ws.write(r, 1, rec.bank, F["txt"])
        ws.write(r, 2, "" if pd.isna(rec.acct) else str(rec.acct), F["txt"])
        ws.write(r, 3, "" if pd.isna(rec.acct_name) else str(rec.acct_name)[:40], F["txt"])
        if pd.notna(rec.date):
            ws.write_datetime(r, 4, pd.Timestamp(rec.date).to_pydatetime(), F["date"])
        ws.write(r, 5, "" if pd.isna(rec.ref) else str(rec.ref)[:40], F["txt"])
        narr = "" if pd.isna(rec.narr) else str(rec.narr)
        ws.write(r, 6, narr[:300], F["txt"])
        ws.write_number(r, 7, float(rec.dr or 0), F["m"])
        ws.write_number(r, 8, float(rec.cr or 0), F["m"])
        ws.write_number(r, 9, float(rec.bal or 0), F["m"])
        ws.write_number(r, 10, int(rec.year), F["i"])
        ws.write(r, 11, rec.category, F["txt"])
        ws.write(r, 12, "" if pd.isna(rec.bucket) else str(rec.bucket)[:40], F["txt"])
        if r % 5000 == 0:
            print("  bank rows", r)
    ws.autofilter(0, 0, len(recs), len(cols) - 1)
    ws.freeze_panes(1, 0)
    ws.set_row(0, 22)

    ws = wb.add_worksheet("Unallocated_1m_plus")
    big = b[b["category"].isin(["UNALLOCATED_IN", "UNALLOCATED_OUT"])].copy()
    big["amt"] = big[["dr", "cr"]].max(axis=1)
    big = big[big["amt"] >= 1_000_000].sort_values("amt", ascending=False)
    for c, h in enumerate(["date", "bank", "direction", "amount", "year", "narration"]):
        ws.write(0, c, h, F["h"])
    ws.set_column(0, 0, 12)
    ws.set_column(1, 1, 14)
    ws.set_column(2, 2, 14)
    ws.set_column(3, 3, 18)
    ws.set_column(4, 4, 10)
    ws.set_column(5, 5, 90)
    for r, rec in enumerate(big.itertuples(index=False), start=1):
        ws.write_datetime(r, 0, pd.Timestamp(rec.date).to_pydatetime(), F["date"])
        ws.write(r, 1, rec.bank, F["txt"])
        ws.write(r, 2, "Receipt" if rec.cr > 0 else "Payment", F["txt"])
        ws.write_number(r, 3, float(rec.amt), F["m"])
        ws.write_number(r, 4, int(rec.year), F["i"])
        ws.write(r, 5, "" if pd.isna(rec.narr) else str(rec.narr)[:400], F["txt"])
    ws.autofilter(0, 0, max(len(big), 1), 5)
    ws.freeze_panes(1, 0)
    print("  unallocated >=1m", len(big))

    ws = wb.add_worksheet("Legend")
    ws.set_column(0, 0, 32)
    ws.set_column(1, 1, 100)
    legend = [
        ("CUSTOMER_COLLECTION", "Bank credit identified as a property receipt, or agreed to a sales-ledger line of the same amount. Turnover is the sales ledger, not this total."),
        ("UNALLOCATED_IN", "Credit the narration does not classify. Liability: receipts pending allocation. Not revenue."),
        ("UNALLOCATED_OUT", "Debit the narration does not classify. Asset: payments pending allocation. Not an expense and not inventory."),
        ("INTERBANK", "Movement with the company's own accounts. Eliminated where it mirrors inside the four accounts; the unmatched net is a financing balance."),
        ("CO_OWNERSHIP_IN / OUT", "Co-ownership investment receipts and returns. Not land or housing turnover."),
        ("OTHER_INCOME_TRADING", "Commodity and other trading receipts, including cashew. Not property turnover."),
        ("INVENTORY_*", "Payments identified as land, housing stock or development materials. Capitalised, then released to cost of sales as revenue is recognised."),
        ("DIRECTOR_REMUNERATION", "Salary and commission narrations naming the directors. Charged to profit. Other director movements are related-party balances."),
        ("DIVIDEND", "Narrations labelled dividend. Distribution, not an expense."),
        ("TAX_PAID", "Narrations that identify a tax payment. Other tax payments may sit in unallocated outflows."),
    ]
    ws.write(0, 0, "Category", F["h"])
    ws.write(0, 1, "How it is used in the revised statements", F["h"])
    for r, (a, b_) in enumerate(legend, start=1):
        ws.write(r, 0, a, F["txt"])
        ws.write(r, 1, b_, F["txt"])
        ws.set_row(r, 32)
    wb.close()
    print("wrote", path)


def write_recon(model):
    path = ROOT / "BAAY_Sales_to_Bank_Reconciliation_Report_revised.xlsx"
    wb = xlsxwriter.Workbook(str(path))
    F = xf(wb)
    b = model["bank"]
    sales = model["sales"]
    ws = wb.add_worksheet("Cover")
    cover(ws, F, "Sales-to-bank reconciliation — revised", [
        "The previous reconciliation used bank inflow totals of about ₦43 million, ₦75 million and ₦131 million. Those are not the totals in the bank extraction. This revision uses the extraction.",
        "Turnover is the sales ledger: 859 lines, ₦2,584,814,000. It is not a percentage of bank inflows and it is not the rounded revenue in the rejected financial statements.",
        "Two different gaps are shown, and they must not be added together. The line match agrees a sales line to one bank credit. The cumulative gap compares the whole sales ledger with every bank credit identified as a property receipt. The financial statements use the cumulative gap.",
        "Co-ownership, service income and credit memos are not in the sales ledger. They are reconciled separately so they are not lost and not double-counted.",
        "Exceptions of ₦138,364,537 were left out of the sales ledger by the source file, mostly because they are aggregates or splits of receipts already counted. They stay out.",
    ])

    ws = wb.add_worksheet("Executive_bridge")
    ws.set_column(0, 0, 62)
    ws.set_column(1, 4, 20)
    headers = ["", "2023", "2024", "2025", "Total"]
    for c, h in enumerate(headers):
        ws.write(0, c, h, F["h"])

    def cat_year(cat, y, side):
        sub = b[(b["year"] == y) & (b["category"] == cat)]
        return float(sub[side].sum())

    rows = []
    # build values
    def add(label, vals, bold=False):
        rows.append((label, vals, bold))

    sales_y = {y: float(sales.loc[sales["year"] == y, "amt"].sum()) for y in (2023, 2024, 2025)}
    add("Sales ledger — customer collections (turnover)", [sales_y[y] for y in (2023, 2024, 2025)], True)
    add("Bank credits identified as property receipts", [cat_year("CUSTOMER_COLLECTION", y, "cr") for y in (2023, 2024, 2025)])
    add("Cumulative sales ledger", [sum(sales_y[y] for y in (2023, 2024, 2025) if y <= t) for t in (2023, 2024, 2025)])
    add("Cumulative property receipts identified", [sum(cat_year("CUSTOMER_COLLECTION", y, "cr") for y in (2023, 2024, 2025) if y <= t) for t in (2023, 2024, 2025)])
    add("Cumulative gap — asset in the financial statements", [model["sfp"][y]["untraced_receipts"] for y in (2023, 2024, 2025)], True)
    add("Gross bank inflows (all credits)", [float(b.loc[b["year"] == y, "cr"].sum()) for y in (2023, 2024, 2025)])
    add("of which inter-account inflows", [cat_year("INTERBANK", y, "cr") for y in (2023, 2024, 2025)])
    add("of which receipts pending allocation", [cat_year("UNALLOCATED_IN", y, "cr") for y in (2023, 2024, 2025)])
    add("of which co-ownership inflows", [cat_year("CO_OWNERSHIP_IN", y, "cr") for y in (2023, 2024, 2025)])
    add("of which trading and other income", [cat_year("OTHER_INCOME_TRADING", y, "cr") + cat_year("OTHER_INCOME_SERVICE", y, "cr") + cat_year("OTHER_INCOME_REFUND", y, "cr") for y in (2023, 2024, 2025)])
    add("Line match — sales lines agreed to one bank credit (count is on the next sheet)", [float(sales.loc[(sales["year"] == y) & sales["bank_txn"].notna(), "amt"].sum()) for y in (2023, 2024, 2025)])
    for r, (label, vals, bold) in enumerate(rows, start=1):
        ws.write(r, 0, label, F["txtb"] if bold else F["txt"])
        fmt = F["mb"] if bold else F["m"]
        for c, v in enumerate(vals):
            ws.write_number(r, 1 + c, v, fmt)
        ws.write_number(r, 4, sum(vals) if "Cumulative" not in label else vals[-1], fmt)
    ws.write(len(rows) + 2, 0, "The cumulative-gap row is the balance sheet asset. Do not add it to the line-match shortfall. They describe the same cash from two angles.", F["warn"])
    ws.set_row(len(rows) + 2, 32)

    ws = wb.add_worksheet("Line_match")
    ws.set_column(0, 0, 55)
    ws.set_column(1, 4, 20)
    for c, h in enumerate(["", "2023", "2024", "2025", "Total"]):
        ws.write(0, c, h, F["h"])
    stats = []
    for y in (2023, 2024, 2025):
        sub = sales[sales["year"] == y]
        stats.append((
            len(sub),
            int(sub["bank_txn"].notna().sum()),
            float(sub.loc[sub["bank_txn"].notna(), "amt"].sum()),
            float(sub.loc[sub["bank_txn"].isna(), "amt"].sum()),
        ))
    labels = [
        ("Sales lines", [s[0] for s in stats], False),
        ("Lines agreed to a bank credit of the same amount", [s[1] for s in stats], False),
        ("Amount agreed line by line", [s[2] for s in stats], True),
        ("Amount not agreed line by line", [s[3] for s in stats], False),
    ]
    for r, (label, vals, bold) in enumerate(labels, start=1):
        ws.write(r, 0, label, F["txtb"] if bold else F["txt"])
        fmt = F["mb"] if bold else (F["i"] if r <= 2 else F["m"])
        for c, v in enumerate(vals):
            ws.write_number(r, 1 + c, v, fmt)
        ws.write_number(r, 4, sum(vals), fmt)
    ws.write(7, 0, "A line can fail this test and still be inside the property-receipt total, because several customer payments are often one bank credit, or the name on the bank is not the name on the ledger. That is why the statements use the cumulative gap of ₦506,439,908 at 31 December 2025, not the larger line-level miss.", F["note"])
    ws.set_row(7, 48)

    ws = wb.add_worksheet("Held_out")
    ws.set_column(0, 0, 55)
    ws.set_column(1, 2, 22)
    ws.write(0, 0, "Population held out of turnover", F["h"])
    ws.write(0, 1, "Amount", F["h"])
    ws.write(0, 2, "Treatment", F["h"])
    held = [
        ("Co-ownership in the exclusions register", 235_707_500, "Liability, not turnover. Bank returns are deducted."),
        ("Service income in the exclusions register", 2_613_000, "Other income where traced; untraced asset where not."),
        ("Credit memos in the exclusions register", 613_250, "Non-cash. Not posted."),
        ("Exceptions tab (secondary records)", model["exceptions_amt"], "Not added. Mostly aggregates or splits of receipts already in the sales ledger."),
        ("2026 client-register rows", 0, "Nine rows. Outside the period. Not in the subscription population."),
    ]
    for r, (a, amt, note) in enumerate(held, start=1):
        ws.write(r, 0, a, F["txt"])
        ws.write_number(r, 1, amt, F["m"])
        ws.write(r, 2, note, F["txt"])
        ws.set_row(r, 28)

    ws = wb.add_worksheet("Conclusion")
    ws.set_column(0, 0, 120)
    lines = [
        "The sales ledger is complete as a turnover record at ₦2,584,814,000. It is on the face of the revised financial statements as customer collections.",
        "The four bank accounts are complete as a cash record. Their stated balances are the cash on the statement of financial position.",
        "The two records do not agree line by line, and they are not forced to agree. The cumulative shortfall of identified property receipts is an asset, described as customer receipts recorded but not covered by identified bank property receipts.",
        "Receipts pending allocation may contain further customer money. They are not reclassified without a name, a keyword or a line match.",
        "This reconciliation replaces the previous sales-to-bank report. The previous bank totals should not be used.",
    ]
    for i, line in enumerate(lines):
        ws.write(i, 0, line, F["note"])
        ws.set_row(i, 36)
    wb.close()
    print("wrote", path)


def write_master(model):
    path = ROOT / "BAAY_PROJECTS_LIMITED_Master_Workbook_2023_2025_revised.xlsx"
    wb = xlsxwriter.Workbook(str(path))
    F = xf(wb)
    ws = wb.add_worksheet("Cover")
    cover(ws, F, "Master workbook 2023–2025 — revised", [
        "This workbook is the accounting model behind the revised financial statements. It is built from the bank extraction, the sales ledger and the client register. It does not use the rounded figures in the previously circulated pack.",
        "Checks: the statement of financial position balances in each year; cash equals the bank; revenue equals collections plus the receivable movement minus the deferred-income movement; the sales ledger totals ₦2,584,814,000.",
        "Profit is stated before allocation of payments the narrations do not identify. Those payments are an asset, scheduled in the bank file. They are not an expense and they are not a plug.",
        "Share capital recognised is ₦1,000,000. The CAC figure of ₦100,000,000 is disclosed and not recognised, because the cash was not traced.",
        "Tax is computed under the law in force before 1 January 2026. The Nigeria Tax Act 2025 does not apply to these years.",
    ])

    ws = wb.add_worksheet("Checks")
    ws.set_column(0, 0, 72)
    ws.set_column(1, 3, 18)
    ws.write(0, 0, "Check", F["h"])
    for i, y in enumerate((2023, 2024, 2025)):
        ws.write(0, 1 + i, y, F["h"])
    checks = []
    for y in (2023, 2024, 2025):
        s = model["sfp"][y]
        p = model["pnl"][y]
        bs = s["total_assets"] - s["total_liabilities"] - s["total_equity"]
        cash = s["cash"] - model["stated_cash"][y]["total"]
        checks.append((bs, cash))
    ws.write(1, 0, "Assets minus liabilities minus equity (must be nil)", F["txt"])
    ws.write(2, 0, "Cash minus stated bank balance (must be nil)", F["txt"])
    ws.write(3, 0, "Result", F["txtb"])
    for i, (bs, cash) in enumerate(checks):
        ws.write_number(1, 1 + i, bs, F["m"])
        ws.write_number(2, 1 + i, cash, F["m"])
        ok = abs(bs) < 1 and abs(cash) < 1
        ws.write(3, 1 + i, "PASS" if ok else "FAIL", F["ok"] if ok else F["bad"])
    ws.write(5, 0, "Sales ledger total", F["txt"])
    ws.write_number(5, 1, float(model["sales"]["amt"].sum()), F["mb"])
    ws.write(6, 0, "Expected", F["txt"])
    ws.write_number(6, 1, 2_584_814_000, F["m"])
    ws.write(7, 0, "IFRS identity: revenue - collections - movement in receivables + movement in deferred income", F["txt"])
    prev_rec, prev_cl = 0.0, 0.0
    for i, y in enumerate((2023, 2024, 2025)):
        s = model["sfp"][y]
        ident = model["pnl"][y]["revenue"] - model["collections"][y] - (s["receivables_gross"] - prev_rec) + (s["contract_liabilities"] - prev_cl)
        ws.write_number(7, 1 + i, ident, F["m"])
        prev_rec, prev_cl = s["receivables_gross"], s["contract_liabilities"]
    ws.write(8, 0, "Identity must be nil. A nil identity means turnover, receivables, deferred income and revenue are the same story.", F["note"])

    def three(ws, title, lines, start=0):
        ws.write(start, 0, title, F["sec"])
        headers = ["", "2023", "2024", "2025"]
        for c, h in enumerate(headers):
            ws.write(start + 1, c, h, F["h"])
        for r, (label, vals, bold) in enumerate(lines):
            ws.write(start + 2 + r, 0, label, F["txtb"] if bold else F["txt"])
            for c, v in enumerate(vals):
                if isinstance(v, str):
                    ws.write(start + 2 + r, 1 + c, v, F["txt"])
                else:
                    ws.write_number(start + 2 + r, 1 + c, float(v), F["mb"] if bold else F["m"])
        return start + 3 + len(lines)

    ws = wb.add_worksheet("SFP")
    ws.set_column(0, 0, 72)
    ws.set_column(1, 3, 20)
    sfp_lines = []
    def grab(key):
        return [model["sfp"][y].get(key, 0) for y in (2023, 2024, 2025)]
    for label, key, bold in [
        ("Property, plant and equipment", "ppe_nbv", False),
        ("Inventories", "inventory", False),
        ("Trade receivables, net", "receivables_net", False),
        ("Customer receipts not covered by identified bank property receipts", "untraced_receipts", False),
        ("Other source receipts not traced to the four banks", "untraced_other", False),
        ("Payments pending allocation", "payments_pending", False),
        ("Foreign-currency and other advances", "fx_advances", False),
        ("Amounts due from related parties", "related_party_dr", False),
        ("Loan repayments in excess of proceeds traced", "borrowings_dr", False),
        ("Cash and cash equivalents", "cash", False),
        ("Total assets", "total_assets", True),
        ("Deferred income — contract liabilities", "contract_liabilities", False),
        ("Co-ownership liabilities", "co_ownership", False),
        ("Receipts pending allocation", "receipts_pending", False),
        ("Net transfers with own accounts outside the extraction", "interaccount_cr", False),
        ("Current tax", "tax_cr", False),
        ("Total liabilities", "total_liabilities", True),
        ("Share capital", "share_capital", False),
        ("Retained earnings", "retained_earnings", False),
        ("Total equity", "total_equity", True),
    ]:
        sfp_lines.append((label, grab(key), bold))
    three(ws, "Statement of financial position at 31 December", sfp_lines)

    ws = wb.add_worksheet("Profit_or_loss")
    ws.set_column(0, 0, 55)
    ws.set_column(1, 3, 20)
    pnl_keys = [
        ("Customer collections — turnover", None),
        ("Revenue — land", "rev_land"),
        ("Revenue — housing", "rev_hous"),
        ("Revenue recognised", "revenue"),
        ("Other income", "other_income"),
        ("Cost of sales", "cos"),
        ("Customer refunds", "refunds"),
        ("Staff costs", "staff"),
        ("Director remuneration", "director_rem"),
        ("Selling commissions", "commission"),
        ("Rent", "rent"),
        ("Professional fees", "professional"),
        ("Administrative expenses", "admin"),
        ("Bank charges (negative is a net credit)", "bank_charges"),
        ("Depreciation", "depreciation"),
        ("Impairment of receivables", "ecl_expense"),
        ("Finance income", "finance_income"),
        ("Finance costs", "finance_cost"),
        ("Profit before tax", "pbt"),
        ("Income tax expense", "tax"),
        ("Profit for the year", "pat"),
        ("Dividends traced (distribution, not an expense)", "dividends"),
    ]
    lines = []
    for label, key in pnl_keys:
        if key is None:
            vals = [model["collections"][y] for y in (2023, 2024, 2025)]
        else:
            vals = [model["pnl"][y].get(key, 0) for y in (2023, 2024, 2025)]
        lines.append((label, vals, key in ("revenue", "pbt", "pat")))
    three(ws, "Profit or loss — and turnover, which is not the same number", lines)
    ws.write(len(lines) + 4, 0, "Profit is before allocation of payments pending allocation. See the sensitivity on the Checks sheet and in the directors' report.", F["warn"])

    ws = wb.add_worksheet("SOCE")
    ws.set_column(0, 0, 42)
    ws.set_column(1, 3, 22)
    for c, h in enumerate(["", "Share capital", "Retained earnings", "Total"]):
        ws.write(0, c, h, F["h"])
    re = model["opening_cash_total"] - model["share_capital"]
    sc = model["share_capital"]
    ws.write(1, 0, "At 1 January 2023", F["txt"])
    ws.write_number(1, 1, sc, F["m"])
    ws.write_number(1, 2, re, F["m"])
    ws.write_number(1, 3, sc + re, F["m"])
    r = 2
    for y in (2023, 2024, 2025):
        pat = model["pnl"][y]["pat"]
        div = model["pnl"][y]["dividends"]
        ws.write(r, 0, f"Profit {y}", F["txt"])
        ws.write_number(r, 2, pat, F["m"])
        ws.write_number(r, 3, pat, F["m"])
        r += 1
        ws.write(r, 0, f"Dividends traced {y}", F["txt"])
        ws.write_number(r, 2, -div, F["m"])
        ws.write_number(r, 3, -div, F["m"])
        re = re + pat - div
        r += 1
        ws.write(r, 0, f"At 31 December {y}", F["txtb"])
        ws.write_number(r, 1, sc, F["mb"])
        ws.write_number(r, 2, re, F["mb"])
        ws.write_number(r, 3, sc + re, F["mb"])
        r += 1

    ws = wb.add_worksheet("Tax")
    ws.set_column(0, 0, 48)
    ws.set_column(1, 3, 20)
    for c, h in enumerate(["", "2023", "2024", "2025"]):
        ws.write(0, c, h, F["h"])
    tax_rows = [
        ("Company size", "size"),
        ("Gross turnover used for the size test", "turnover"),
        ("Rate", "rate"),
        ("Profit before tax", "pbt"),
        ("Assessable profit", "assessable"),
        ("Companies income tax", "cit_payable"),
        ("Tertiary education tax at 3%", "tet"),
        ("Police trust fund levy", "ptf"),
        ("NASENI levy (confirm sector scope)", "naseni"),
        ("Current tax expense", "current"),
        ("Minimum tax applied", "min_applied"),
    ]
    for r, (label, key) in enumerate(tax_rows, start=1):
        ws.write(r, 0, label, F["txt"])
        for c, y in enumerate((2023, 2024, 2025)):
            v = model["tax"][y][key]
            if isinstance(v, str):
                ws.write(r, 1 + c, v, F["txt"])
            elif isinstance(v, bool):
                ws.write(r, 1 + c, "Yes" if v else "No", F["txt"])
            elif key == "rate":
                ws.write(r, 1 + c, f"{v:.0%}", F["txt"])
            else:
                ws.write_number(r, 1 + c, float(v), F["m"])
    ws.write(14, 0, "Years ended on or before 31 December 2025 stay under the repealed law. The 4% development levy in the Nigeria Tax Act 2025 starts with periods from 1 January 2026 and is not accrued.", F["note"])
    ws.set_row(14, 32)

    ws = wb.add_worksheet("Contracts")
    c = model["contracts"]
    cols = ["ref", "name", "product", "project", "status", "contract", "policy",
            "col_2023", "col_2024", "col_2025", "rev_2023", "rev_2024", "rev_2025",
            "rec_2025", "cl_2025", "revenue_total"]
    for i, h in enumerate(cols):
        ws.write(0, i, h, F["h"])
        ws.set_column(i, i, 18 if h not in ("customer", "policy", "project") else 36)
    ws.set_column(6, 6, 55)
    for r, rec in enumerate(c[cols].itertuples(index=False), start=1):
        for i, val in enumerate(rec):
            if pd.isna(val):
                ws.write(r, i, "", F["txt"])
            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                ws.write_number(r, i, float(val), F["m"])
            else:
                ws.write(r, i, str(val)[:180], F["txt"])
    ws.autofilter(0, 0, len(c), len(cols) - 1)
    ws.freeze_panes(1, 2)

    ws = wb.add_worksheet("Comparison")
    ws.set_column(0, 0, 28)
    ws.set_column(1, 6, 18)
    headers = ["Year", "Old revenue", "Revised collections", "Revised revenue", "Old cash", "Revised cash", "Old deferred", "Revised deferred"]
    for i, h in enumerate(headers):
        ws.write(0, i, h, F["h"])
    for r, y in enumerate((2023, 2024, 2025), start=1):
        ws.write(r, 0, y, F["txt"])
        ws.write_number(r, 1, OLD[y]["revenue"], F["m"])
        ws.write_number(r, 2, model["collections"][y], F["mb"])
        ws.write_number(r, 3, model["pnl"][y]["revenue"], F["m"])
        ws.write_number(r, 4, OLD[y]["cash"], F["m"])
        ws.write_number(r, 5, model["sfp"][y]["cash"], F["mb"])
        ws.write_number(r, 6, OLD[y]["deferred"], F["m"])
        ws.write_number(r, 7, model["sfp"][y]["contract_liabilities"], F["mb"])
    ws.write(5, 0, "The old figures are the pack that was circulated and rejected. They do not agree to the sales ledger or the bank. They are shown only so the difference is visible.", F["warn"])
    ws.set_row(5, 32)

    ws = wb.add_worksheet("Policies")
    ws.set_column(0, 0, 120)
    policies = [
        "Land revenue: bought, closed or completed, starting from contract value. The receivable is capped at the client-register balance, and revenue is reduced by the same amount (₦515,131,000 across the three years). Paying, promise, refund and migrated stay deferred.",
        "Housing revenue: only on evidenced handover (Bought or Completed, or Paid with register paid-to-date of at least 80% of contract). Other housing receipts are deferred income.",
        "A receivable is the unpaid balance of a recognised contract only. Unpaid balances on deferred contracts are a pipeline, disclosed, not a receivable.",
        "Loss allowance: 5% general. No loss history was in the files.",
        "Inventory: narration-identified land, housing stock and development cost, released in proportion to revenue over revenue plus deferred income of the same project.",
        "Opening balances at 1 January 2023: cash only, ₦478,080.25. No opening land, receivable or payable was supported.",
        "Share capital recognised: ₦1,000,000. CAC issued capital of ₦100,000,000 is disclosed, not recognised.",
        "Pending allocation is not forced into revenue, expense or inventory.",
        "Residential plot and housing sales are treated as outside VAT. A contrary view by the tax authority is a contingent exposure, not an accrual.",
        "No audit-fee accrual was booked. The fee is for the engagement partner to accrue on completion.",
    ]
    for i, line in enumerate(policies):
        ws.write(i, 0, line, F["note"])
        ws.set_row(i, 28)
    wb.close()
    print("wrote", path)


def write_sales(model):
    src = ROOT / "BAAY_Sales_and_Customers_2022_2025_for_Audit.xlsx"
    path = ROOT / "BAAY_Sales_and_Customers_2022_2025_for_Audit_revised.xlsx"
    shutil.copy(src, path)
    wb = load_workbook(path)
    navy = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    fill = PatternFill("solid", fgColor="1F4E79")
    title_font = Font(name="Calibri", bold=True, color="1F4E79", size=16)
    note_font = Font(name="Calibri", size=10)
    thin = Border(
        left=Side(style="thin", color="D0D7E2"),
        right=Side(style="thin", color="D0D7E2"),
        top=Side(style="thin", color="D0D7E2"),
        bottom=Side(style="thin", color="D0D7E2"),
    )
    money = "₦#,##0.00"

    ws = wb.create_sheet("00_Revision_cover", 0)
    ws["A1"] = "BAAY PROJECTS LIMITED  ·  RC 1526224"
    ws["A1"].font = Font(name="Calibri", bold=True, color="1F4E79", size=12)
    ws["A2"] = "Sales and customers — revised"
    ws["A2"].font = title_font
    ws["A3"] = "REVISED  ·  8 October 2026  ·  original sheets retained so their formulas still work"
    ws["A3"].font = Font(name="Calibri", bold=True, color="8C2F2F", size=10)
    notes = [
        "The original sheets are unchanged. The revision is the two sheets added at the front: this cover, and IFRS15_recognition.",
        "Turnover remains the Sales_Ledger: 859 lines, ₦2,584,814,000 for 2023 to 2025. Nothing was added from the Exceptions tab, and nothing was scaled.",
        "IFRS15_recognition applies the revenue policy to each subscription and to uncontracted sales-ledger lines. Revenue, the receivable and the deferred income on that sheet are the figures in the revised financial statements.",
        "Land marked bought or closed starts from the contract value. The receivable is capped at the client-register balance, and revenue on the sheet is reduced by the same amount. That reduction is ₦515,131,000 across 2023 to 2025. If the register is incomplete, revenue and receivables are both understated. Housing is recognised only on evidenced handover. Other collections are deferred income.",
        "A receivable is only the unpaid balance of a recognised contract, and not above the register balance. Unpaid balances on contracts still deferred are not receivables. They are a pipeline, disclosed in the financial statements.",
        "Exclusions (co-ownership, service income, credit memos) stay off turnover. Co-ownership is a liability in the revised statements. Credit memos are non-cash and are not posted.",
        "Client-register rows dated 2026 are outside this revision.",
    ]
    for i, text in enumerate(notes):
        ws.cell(5 + i, 1, text).font = note_font
        ws.row_dimensions[5 + i].height = 32
    ws.column_dimensions["A"].width = 140

    ws = wb.create_sheet("IFRS15_recognition", 1)
    c = model["contracts"]
    headers = ["Sub ref", "Customer", "Product", "Project", "Status", "Contract value",
               "Policy", "Collections 2023", "Collections 2024", "Collections 2025",
               "Revenue 2023", "Revenue 2024", "Revenue 2025",
               "Receivable at 31 Dec 2025", "Deferred income at 31 Dec 2025"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(1, col, h)
        cell.font = navy
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 30
    ws.auto_filter.ref = f"A1:O{len(c)+1}"
    ws.freeze_panes = "C2"
    widths = [16, 32, 12, 28, 18, 18, 55, 16, 16, 16, 16, 16, 16, 22, 28]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    money_cols = set(range(6, 7)) | set(range(8, 16))
    for r, rec in enumerate(c.itertuples(index=False), start=2):
        values = [
            rec.ref, rec.name, rec.product, rec.project, rec.status, rec.contract, rec.policy,
            rec.col_2023, rec.col_2024, rec.col_2025, rec.rev_2023, rec.rev_2024, rec.rev_2025,
            rec.rec_2025, rec.cl_2025,
        ]
        for col, val in enumerate(values, start=1):
            cell = ws.cell(r, col, None if pd.isna(val) else val)
            cell.font = Font(name="Calibri", size=9)
            cell.border = thin
            cell.alignment = Alignment(vertical="center", wrap_text=(col == 7))
            if col in money_cols and isinstance(val, (int, float)) and not isinstance(val, bool):
                cell.number_format = money
    # totals
    total_row = len(c) + 2
    ws.cell(total_row, 1, "Total").font = Font(name="Calibri", bold=True, size=9)
    for col, attr in [(6, "contract"), (8, "col_2023"), (9, "col_2024"), (10, "col_2025"),
                      (11, "rev_2023"), (12, "rev_2024"), (13, "rev_2025"), (14, "rec_2025"), (15, "cl_2025")]:
        cell = ws.cell(total_row, col, float(c[attr].sum()))
        cell.font = Font(name="Calibri", bold=True, size=9)
        cell.number_format = money
        cell.fill = PatternFill("solid", fgColor="E7EEF6")
    wb.save(path)
    print("wrote", path)


def main():
    print("loading model")
    model = pickle.load(open(MODEL, "rb"))
    write_recon(model)
    write_master(model)
    write_sales(model)
    write_bank(model)


if __name__ == "__main__":
    main()
