"""
Rebuild Baay Projects Limited's 2023-2025 financial statements from the
source workbooks on this branch.

Cash comes from the four extracted bank accounts. Customer turnover,
receivables and deferred income come from the sales ledger and the client
register, bridged to the bank. Nothing is plugged into revenue.

Run:  /tmp/venv/bin/python revised/engine.py
Writes: /tmp/baay_extract/model.pkl
"""

from __future__ import annotations

import pickle
import re
from collections import defaultdict
from datetime import datetime

import pandas as pd

BANK_PATH = "Baay_Projects_Clean_Bank_Statements_Extraction_2023_2025.xlsx"
SALES_PATH = "BAAY_Sales_and_Customers_2022_2025_for_Audit.xlsx"

OPENING_CASH = {
    "First Bank": 12971.91,
    "GTBank": 0.0,
    "Sterling": 944.19,
    "Providus": 464164.15,
}
OPENING_CASH_TOTAL = round(sum(OPENING_CASH.values()), 2)  # 478,080.25
SHARE_CAPITAL = 1_000_000.00

# CAC status report 26 Aug 2026 — disclosed, not recognised as paid-in cash.
CAC_ISSUED_CAPITAL = 100_000_000.00

TITLES = {
    "MR", "MRS", "MS", "MISS", "DR", "ENGR", "CHIEF", "ALH", "ALHAJI", "ALHAJA",
    "PASTOR", "BARR", "BARRISTER", "PRINCE", "PRINCESS", "SIR", "HON", "REV",
    "PROF", "ARC", "ARCH", "ENGR.", "MR.", "MRS.", "OTUNBA", "CHIEF.",
}

DIRECTOR_NS = [
    "ADEGOKESEGUNBABATUNDE",
    "SEGUNBABATUNDE",
    "ADEGOKESEGUN",
    "MARYAYOBOADE",
    "AYOBOADEADEGOKE",
    "ADEGOKE DENISA".replace(" ", ""),
    "ADEGOKE RAHEEM",
    "ADEGOKE RAHEEMADEBAYO",
    "AJIBOLAOLUWATOBI",
    "SHUAIBSULIAT",
    "KESHINROPHEBE",
    "OWOLABICHARLES",
    "OLAYINKAOLADOTUN",
    "NOAHABDULAZEEZ",
]
# keep the spaced ones compacted
DIRECTOR_NS = [re.sub(r"[^A-Z0-9]", "", x) for x in [
    "ADEGOKE SEGUN BABATUNDE",
    "ADEGOKE SEGUN BABA",
    "ADEGOKE SEGUN",
    "SEGUN ADEGOKE",
    "SEGUN BABATUNDE",
    "MARY AYOBOADE",
    "AYOBOADE ADEGOKE",
    "ADEGOKE DENISA",
    "ADEGOKE RAHEEM",
    "AJIBOLA OLUWATOBI",
    "SHUAIB SULIAT",
    "KESHINRO PHEBE",
    "OWOLABI CHARLES OLUWATOBI",
    "OLAYINKA OLADOTUN",
    "NOAH ABDULAZEEZ",
]]

# Counterparties that the client register shows as customers, realtors or
# payment agents. Credits from these names are property collections.
KNOWN_COLLECTORS = [
    "LANDNEST", "SALEMREALESTATE", "SCRAPAYS", "SCRAPAYSTECHNOLOG",
    "EBOSELE", "OGBOMAH", "AKINJOGUNLA", "OMOKOYA", "LAWSONALBERT",
    "AKINSANYA", "AKINSANYA BIMPE".replace(" ", ""),
]

VENDOR_INVENTORY = [
    ("QCHOMES", "GVC", "housing"),
    ("MAYNCERY", "Other land", "land"),
    ("TITILOYE", "Other land", "land"),
    ("OSUNLAJA", "Green City", "land"),
    ("POPOOLAOLAYEMI", "Other land", "land"),
    ("POPOOLAOLAYEMIHAKEEM", "Other land", "land"),
]


def ns(text) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(text).upper())


def tokens(name) -> list[str]:
    s = re.sub(r"[^A-Z ]", " ", str(name).upper())
    return [t for t in s.split() if t not in TITLES and len(t) >= 4]


def parse_date(value):
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, datetime):
        return pd.Timestamp(value)
    if isinstance(value, pd.Timestamp):
        return value
    s = str(value).strip()
    for fmt in ("%d-%b-%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%b-%y"):
        try:
            return pd.Timestamp(datetime.strptime(s[:11] if fmt == "%d-%b-%Y" else s, fmt))
        except Exception:
            continue
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def money(x) -> float:
    return round(float(x) + 1e-9, 2)


# ---------------------------------------------------------------------------
# Load
# ---------------------------------------------------------------------------

def load_bank() -> pd.DataFrame:
    frames = []
    g = pd.read_excel(BANK_PATH, sheet_name="GTBank - Transactions", header=7)
    g = g[pd.to_numeric(g["Line #"], errors="coerce").notna()]
    frames.append(pd.DataFrame({
        "bank": "GTBank", "acct": "0893079582", "acct_name": "BAAY PROJECTS LTD",
        "date": g["Trans Date"].map(parse_date),
        "narr": g["Transaction Remarks / Description"].astype(str),
        "ref": g["Reference"].astype(str),
        "dr": pd.to_numeric(g["Debit (DR)"], errors="coerce").fillna(0),
        "cr": pd.to_numeric(g["Credit (CR)"], errors="coerce").fillna(0),
        "bal": pd.to_numeric(g["Stated Balance"], errors="coerce"),
        "line": g["Line #"],
    }))
    f = pd.read_excel(BANK_PATH, sheet_name="First Bank - Transactions", header=7)
    f = f[pd.to_numeric(f["Line #"], errors="coerce").notna()]
    frames.append(pd.DataFrame({
        "bank": "First Bank", "acct": "2033736938", "acct_name": "BAAY DEGOK NIG LTD",
        "date": f["Trans Date"].map(parse_date),
        "narr": f["Transaction Details / Description"].astype(str),
        "ref": f["Ref Number"].astype(str),
        "dr": pd.to_numeric(f["Withdrawal (DR)"], errors="coerce").fillna(0),
        "cr": pd.to_numeric(f["Deposit (CR)"], errors="coerce").fillna(0),
        "bal": pd.to_numeric(f["Stated Balance"], errors="coerce"),
        "line": f["Line #"],
    }))
    s = pd.read_excel(BANK_PATH, sheet_name="Sterling Bank - Transactions", header=7)
    s = s[pd.to_numeric(s["Line #"], errors="coerce").notna()]
    frames.append(pd.DataFrame({
        "bank": "Sterling", "acct": "0091166190", "acct_name": "BAAY PROJECTS",
        "date": s["Trans Date"].map(parse_date),
        "narr": s["Narration / Transaction Description"].astype(str),
        "ref": s["Doc Ref"].astype(str),
        "dr": pd.to_numeric(s["Money Out / Debit (DR)"], errors="coerce").fillna(0),
        "cr": pd.to_numeric(s["Money In / Credit (CR)"], errors="coerce").fillna(0),
        "bal": pd.to_numeric(s["Stated Balance"], errors="coerce"),
        "line": s["Line #"],
    }))
    p = pd.read_excel(BANK_PATH, sheet_name="Providus - Transactions", header=7)
    p = p[pd.to_numeric(p["Line #"], errors="coerce").notna()]
    frames.append(pd.DataFrame({
        "bank": "Providus", "acct": "5400724970", "acct_name": "BAAY PROJECTS LTD",
        "date": p["Trans Date"].map(parse_date),
        "narr": p["Remarks / Narration"].astype(str),
        "ref": "",
        "dr": pd.to_numeric(p["Debit (DR)"], errors="coerce").fillna(0),
        "cr": pd.to_numeric(p["Credit (CR)"], errors="coerce").fillna(0),
        "bal": pd.to_numeric(p["Stated Balance"], errors="coerce"),
        "line": p["Line #"],
    }))
    b = pd.concat(frames, ignore_index=True)
    b["dr"] = b["dr"].map(money)
    b["cr"] = b["cr"].map(money)
    b["ns"] = b["narr"].map(ns)
    b["year"] = b["date"].dt.year
    # Drop opening/closing marker rows (no date, or zero movement labelled opening).
    b = b[b["date"].notna() & b["year"].between(2023, 2025)].copy()
    b = b[~b["ns"].str.contains("OPENINGBALANCE")].copy()
    b = b.reset_index(drop=True)
    b["txn_id"] = b.index.astype(int)
    return b


def load_sales_pack():
    sales = pd.read_excel(SALES_PATH, sheet_name="Sales_Ledger", header=3)
    subs = pd.read_excel(SALES_PATH, sheet_name="Subscriptions", header=3)
    cust = pd.read_excel(SALES_PATH, sheet_name="Customer_Master", header=3)
    excl = pd.read_excel(SALES_PATH, sheet_name="Exclusions", header=3)
    exc = pd.read_excel(SALES_PATH, sheet_name="Exceptions_Dec23_Dec24", header=3)

    sales = sales[sales["Ref"].notna() & ~sales["Ref"].astype(str).str.upper().str.contains("TOTAL")].copy()
    subs = subs[subs["Sub ref"].notna() & ~subs["Sub ref"].astype(str).str.upper().str.contains("TOTAL")].copy()
    cust = cust[cust["Customer ID"].notna()].copy()
    excl = excl[excl["Ref"].notna() & ~excl["Ref"].astype(str).str.upper().str.contains("TOTAL")].copy()
    exc = exc[exc["Ref"].notna() & ~exc["Ref"].astype(str).str.upper().str.contains("TOTAL")].copy()

    sales["amt"] = pd.to_numeric(sales["Amount (₦)"], errors="coerce").fillna(0).map(money)
    sales["date"] = pd.to_datetime(sales["Receipt date"], errors="coerce")
    sales["year"] = sales["Year"].astype(int)
    for c in ["Contract value (₦)", "Paid to date per register (₦)", "Balance per register (₦)"]:
        subs[c] = pd.to_numeric(subs[c], errors="coerce")
    excl["amt"] = pd.to_numeric(excl["Amount (₦)"], errors="coerce").fillna(0).map(money)
    excl["date"] = pd.to_datetime(excl["Date"], errors="coerce")
    excl["year"] = pd.to_numeric(excl["Year"], errors="coerce")
    exc["amt"] = pd.to_numeric(exc["Amount (₦)"], errors="coerce").fillna(0).map(money)
    return sales, subs, cust, excl, exc


# ---------------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------------

def inflow_interbank(row_ns: str) -> bool:
    if "ADEGOKE" in row_ns and "BAAYPROJECTSLTD" not in row_ns and "BAAYDEGOK" not in row_ns:
        return False
    if "SENDERBAAYPROJECTS" in row_ns or "SENDERBAAYDEGOK" in row_ns:
        return True
    if "TRANSFERFROMBAAYPROJECTS" in row_ns or "TRANSFERFROMBAAYDEGOK" in row_ns:
        return True
    if "TRFFRMBAAYPROJECTS" in row_ns or "TRFFRMBAAYDEGOK" in row_ns:
        return True
    if "FTFROMBAAYPROJECTS" in row_ns or "FTFROMBAAYDEGOK" in row_ns:
        return True
    if "FROMBAAYPROJECTSLTD" in row_ns or "FROMBAAYPROJECTSLIMITED" in row_ns:
        return True
    if "FROMBAAYDEGOKNIG" in row_ns or "FROMBAAYDEGOK" in row_ns:
        return True
    if "FBNMOBILEBAAYPROJECTSLTD" in row_ns and "ADEGOKE" not in row_ns:
        return True
    # Receiving leg of an own-account internet-banking transfer.
    if re.search(r"FIP(?:PRB|STL|FBN|GTB|ACC)?BAAYPROJECTS", row_ns) and "ADEGOKE" not in row_ns:
        return True
    if "IFOROEBAAYPROJECTS" in row_ns or "IFOBAAYPROJECTS" in row_ns:
        return True
    if ("BAAYDEGOK" in row_ns or "BAAYPROJECTSLTD" in row_ns) and "ADEGOKE" not in row_ns:
        if any(k in row_ns for k in ("CIBCASH", "INWARDTRANSFER", "CLEARINGDEPOSIT", "FROMFBN", "FROMSTERLING")):
            return True
    return False


def outflow_interbank(row_ns: str) -> bool:
    if any(v in row_ns for v in ("QCHOMES", "MAYNCERY", "WITFORD", "TITILOYE", "POPOOLA", "ADEGOKE", "ARUBI")):
        # beneficiary is a third party or a director, even if the sender name is also present
        if "BAAYPROJECTSLTDCIB" not in row_ns and "BAAYDEGOKNIGLTDCIB" not in row_ns:
            if "IFOBAAY" not in row_ns and "TOBAAYPROJECTSLTD" not in row_ns and "TOBAAYDEGOK" not in row_ns:
                return False
    # Internet banking to another account of the company: FIP:IB:PRB/BAAY PROJECTS LTD/CIB:...
    if "BAAYPROJECTSLTDCIB" in row_ns or "BAAYDEGOKNIGLTDCIB" in row_ns or "BAAYDEGOKCIB" in row_ns:
        return True
    if re.search(r"FIPIB(?:PRB|STL|FBN|GTB|ACC)?BAAYPROJECTS", row_ns) and "ADEGOKE" not in row_ns:
        return True
    if re.search(r"(IFO|TO)BAAYPROJECTS(LTD|LIMITED|NIG|NIGERIA)", row_ns):
        return True
    if re.search(r"(IFO|TO)BAAYDEGOK", row_ns):
        return True
    if "TOSTERLINGBANKBAAYPROJECTS" in row_ns and "QCHOMES" not in row_ns:
        return True
    if "TRANSFERTOBAAYPROJECTS" in row_ns or "TRANSFERTOBAAYDEGOK" in row_ns:
        return True
    if re.search(r"TOBAAYPROJECTS[0-9]", row_ns) and "QCHOMES" not in row_ns:
        return True
    if row_ns.endswith("TOBAAY") or "TRANSFERFROMBAAYPROJECTSTOBA" in row_ns:
        return True
    return False


def is_director(row_ns: str) -> bool:
    return any(d in row_ns and len(d) >= 12 for d in DIRECTOR_NS)


def property_keyword(row_ns: str) -> bool:
    keys = (
        "GREENCITY", "PACIFICAPARTMENT", "PACIFICCOURT", "GORGEVIEW", "GVC",
        "FORESHORE", "HERITAGEGARDEN", "CORNERPIECE", "CORNERPIECE",
        "DEVELOPMENTALFEE", "DEVELOPMENTFEE", "DOCUMENTATIONFEE",
        "PLOTPAYMENT", "INITIALDEPOSIT", "SQM", "KETUEPE", "LANTSALE",
        "PLOTPRICE", "ALLOCATION", "INSTALMENT", "INSTALLMENT",
        "PACIFICAPARTMENTS", "UNIT8", "ATICO", "ATICAN",
    )
    return any(k in row_ns for k in keys)


def trading_keyword(row_ns: str) -> bool:
    return any(k in row_ns for k in (
        "CASHEW", "RAWCASHEW", "FISHMONGER", "GOODSSOLD", "PAYMENTOFGOODS",
        "27TONS", "CASHEWNUT",
    ))


def bank_charge(row_ns: str, amount: float, direction: str) -> bool:
    if amount >= 5000 and "ACCOUNTMAINTENANCE" not in row_ns and "AMF" not in row_ns:
        # large amounts are not routine bank charges, even if VAT is mentioned
        if not any(k in row_ns for k in ("ACCOUNTMAINTENANCE", "INTERESTCAPITAL")):
            if amount >= 20000:
                return False
    keys = (
        "STAMPDUTY", "EMTLEVY", "NIPFEE", "NIPVAT", "SMSALERT", "SMSCHARGE",
        "ACCOUNTMAINTENANCE", "FIPCHARGES", "ELECTRONICMONEYTRANSFERLEVY",
        "ELECMONEYTRSFLEVY", "VATFEE", "NIPCHARGE", "TRANSFERCHARGE",
        "COTCHARGE", "CARDMAINTENANCE", "MONTHLYMAINTENANCE",
        "GOVTLEVY", "GOVTLEVYFIRSEMT",
    )
    if any(k in row_ns for k in keys):
        return True
    if "VAT" in row_ns and amount < 5000 and direction == "dr":
        # VAT on a transfer fee, not a VAT remittance
        if any(k in row_ns for k in ("NIP", "FEE", "CHARGE", "TRANSFER", "COMMISSION")):
            return True
    if "COMMISSION" in row_ns and amount < 1000:
        return True
    if row_ns.endswith("VAT") and amount < 2000:
        return True
    return False


def classify_row(row_ns: str, dr: float, cr: float) -> tuple[str, str]:
    """Return (category, project_bucket)."""
    direction = "cr" if cr > dr else "dr"
    amount = cr if direction == "cr" else dr

    if "REVERSAL" in row_ns and direction == "cr" and amount < 100_000:
        return "BANK_CHARGE_CONTRA", ""
    if bank_charge(row_ns, amount, direction):
        if "INTEREST" in row_ns and direction == "dr":
            return "FINANCE_COST", ""
        if "INTEREST" in row_ns and direction == "cr":
            return "FINANCE_INCOME", ""
        return "BANK_CHARGE", ""
    if "INTERESTCAPITAL" in row_ns or (row_ns == "INTERESTCAPITALISED"):
        return ("FINANCE_COST", "") if direction == "dr" else ("FINANCE_INCOME", "")

    if direction == "cr" and inflow_interbank(row_ns):
        return "INTERBANK", ""
    if direction == "dr" and outflow_interbank(row_ns):
        return "INTERBANK", ""

    # Known inventory vendors (debits)
    if direction == "dr":
        for key, bucket, kind in VENDOR_INVENTORY:
            if key in row_ns:
                return ("INVENTORY_HOUSING" if kind == "housing" else "INVENTORY_LAND"), bucket

    # Trading
    if "WITFORD" in row_ns and direction == "cr":
        return "OTHER_INCOME_TRADING", ""
    if "GASLIFT" in row_ns or "GASLIFT" in row_ns.replace("GASL", "GASLIFT"):
        pass
    if "GASLIFT" in row_ns and direction == "cr":
        return "OTHER_INCOME_REFUND", ""
    if "RAOFISH" in row_ns or "RAOFSIH" in row_ns or "RAOFSI" in row_ns:
        if direction == "cr" and property_keyword(row_ns):
            return "CUSTOMER_COLLECTION", "Pacific"
        if direction == "cr":
            return "OTHER_INCOME_TRADING", ""
        if direction == "dr":
            return "TRADING_PURCHASES", ""

    # Co-ownership
    if any(k in row_ns for k in ("COOWNER", "COOWNERSHIP", "FIXEDRETURN")):
        return ("CO_OWNERSHIP_IN" if direction == "cr" else "CO_OWNERSHIP_OUT"), ""

    # Tax — whole-token, never the substring inside FIRST
    if re.search(r"(PAYE|LIRS|WITHHOLDINGTAX|WHTREMIT|FIRSLEVY|TAXPAYMENT|CITPAYMENT)", row_ns):
        if "FIRSTBANK" not in row_ns or any(k in row_ns for k in ("PAYE", "LIRS", "WITHHOLDING", "TAXPAY")):
            if any(k in row_ns for k in ("PAYE", "LIRS", "WITHHOLDINGTAX", "WHTREMIT")):
                return "TAX_PAID", ""

    # Director salary / commission before generic director current account
    director = is_director(row_ns)
    if direction == "dr" and "SALARY" in row_ns:
        if director:
            return "DIRECTOR_REMUNERATION", ""
        return "STAFF_COST", ""
    if direction == "dr" and "COMMISSION" in row_ns and amount >= 1000:
        return "SELLING_COMMISSION", ""
    if direction == "dr" and any(k in row_ns for k in ("PAYROLL", "WAGES", "STAFFSALARY")):
        return "STAFF_COST", ""

    # Refunds
    if "REFUND" in row_ns:
        return ("OTHER_INCOME_REFUND" if direction == "cr" else "CUSTOMER_REFUND"), ""

    # Property collections (credits)
    if direction == "cr":
        if any(k in row_ns for k in KNOWN_COLLECTORS):
            return "CUSTOMER_COLLECTION", ""
        if property_keyword(row_ns) and not trading_keyword(row_ns):
            return "CUSTOMER_COLLECTION", ""
        if trading_keyword(row_ns):
            return "OTHER_INCOME_TRADING", ""

    # Inventory / development (debits)
    if direction == "dr":
        if any(k in row_ns for k in (
            "CEMENT", "GRANITE", "IRONROD", "REINFORCEMENT", "DANGOTE",
            "ROOFING", "PLUMBING", "ALUMINIUM", "ALUMINUM", "HARDCORE",
            "LATERITE", "EXCAVAT", "BULLDOZ", "MARSHAL", "BLOCKS",
            "BUILDINGMATERIAL", "WORKMANSHIP", "MASON", "CARPENTER",
            "TILER", "WELDER", "BRICKLAYER", "SITEEXPENSE",
        )):
            return "INVENTORY_DEVELOPMENT", "General development"
        if re.search(r"(SAND|PAINT|TIMBER|PLANK|WOOD|TILES|WINDOW|DOOR)", row_ns) and amount >= 20000:
            return "INVENTORY_DEVELOPMENT", "General development"
        if any(k in row_ns for k in ("OMOONILE", "EXCISION", "LANDOWNER", "LANDPAYMENT", "LANDPURCHASE")):
            return "INVENTORY_LAND", "Other land"
        if "SURVEY" in row_ns and "EQUIPMENT" not in row_ns:
            return "INVENTORY_LAND", "Other land"
        if "DEED" in row_ns and amount >= 20000:
            return "INVENTORY_LAND", "Other land"
        if "GREENCITY" in row_ns or "KETUEPE" in row_ns or "LANTABA" in row_ns:
            return "INVENTORY_LAND", "Green City"
        if "QCHOMES" in row_ns or "GORGEVIEW" in row_ns:
            return "INVENTORY_HOUSING", "GVC"

    # Operating expenses
    if direction == "dr":
        if "RENT" in row_ns and amount >= 20000:
            return "RENT", ""
        if any(k in row_ns for k in ("LEGAL", "LAWYER", "SOLICITOR", "BARRISTER", "AUDITFEE")):
            return "PROFESSIONAL", ""
        if any(k in row_ns for k in (
            "FUEL", "DIESEL", "PMS", "TRANSPORT", "LOGISTICS", "AIRTIME",
            "DSTV", "GOTV", "PHCN", "IKEDC", "EKEDC", "ELECTRICITY",
            "STATIONERY", "PRINTING", "HOTEL", "FLIGHT", "AIRPEACE",
            "WELFARE", "ALLOWANCE", "SECURITY", "ENTERTAINMENT",
            "HOSPITAL", "PHARMACY", "HMO",
        )):
            return "ADMIN_OPS", ""
        if ("POS" in row_ns or "WEBPURCHASE" in row_ns or "WEBPurchase".upper() in row_ns) and amount < 2_000_000:
            return "ADMIN_OPS", ""
        # PPE
        if amount >= 100_000 and any(k in row_ns for k in (
            "TOYOTA", "HILUX", "LEXUS", "LAPTOP", "GENERATOR", "FURNITURE",
            "AIRCONDITION", "COMPUTER", "PRINTER", "EQUIPMENT",
        )):
            if "FUEL" not in row_ns and "DIESEL" not in row_ns:
                return "PPE", ""

    # Commodity purchases and land-vendor payments that the narration names.
    if direction == "dr" and "CASHEW" in row_ns:
        return "TRADING_PURCHASES", ""
    if direction == "dr" and any(k in row_ns for k in ("ARUBIEWE", "ARUBI", "ABOGUN", "ATICAN", "TINATH", "BARUWA", "AJODANEWTOWN")):
        return "INVENTORY_LAND", "Other land"
    if direction == "dr" and "REPAY" in row_ns and amount >= 50_000:
        return "BORROWINGS_OUT", ""
    if direction == "dr" and "LEMONADE" in row_ns and ("ADEGOKE" in row_ns or "SEGUNADEGOKE" in row_ns):
        return "RELATED_PARTY_OUT", ""
    if "DIVIDEND" in row_ns and direction == "dr":
        return "DIVIDEND", ""
    if direction == "cr" and ("FINCRA" in row_ns or "BUILDINGPROJECT" in row_ns or "TRFBUILDINGPROJECT" in row_ns):
        return "CUSTOMER_COLLECTION", ""
    if direction == "cr" and "TITILOYE" in row_ns:
        return "INVENTORY_CONTRA", ""
    if direction == "cr" and "PAYMENTOFGOODS" in row_ns and "AKINSOLA" not in row_ns:
        return "OTHER_INCOME_TRADING", ""

    # Director / related party residual
    if director and amount >= 1000:
        if direction == "dr" and any(k in row_ns for k in ("DIVIDEND", "SHARE")):
            return "DIVIDEND", ""
        return ("RELATED_PARTY_IN" if direction == "cr" else "RELATED_PARTY_OUT"), ""

    # FX
    if direction == "dr" and any(k in row_ns for k in ("EXCHANGEPOUNDS", "EXCHANGEDOLLARS", "FXPURCHASE", "DOLLAREXCHANGE")):
        return "FX_ADVANCE", ""
    if "VAULTA" in row_ns and direction == "dr":
        return "FX_ADVANCE", ""

    # Loans
    if "LOAN" in row_ns and amount >= 100_000:
        if direction == "cr":
            return "BORROWINGS_IN", ""
        if any(k in row_ns for k in ("REPAY", "REPAYMENT", "MFB", "MICROFINANCE")):
            return "BORROWINGS_OUT", ""

    if direction == "cr":
        return "UNALLOCATED_IN", ""
    return "UNALLOCATED_OUT", ""


def apply_name_overrides(bank: pd.DataFrame, customers: list[dict]) -> None:
    """Credits not otherwise specific: match a customer / realtor name."""
    # Only reclassify residual or generic credits. Do not override interbank,
    # bank charges, trading, inventory.
    locked = {
        "INTERBANK", "BANK_CHARGE", "BANK_CHARGE_CONTRA", "FINANCE_COST",
        "FINANCE_INCOME", "OTHER_INCOME_TRADING", "OTHER_INCOME_REFUND",
        "CO_OWNERSHIP_IN", "TAX_PAID",
    }
    # Build search list: nospace name >= 10
    names = []
    for c in customers:
        compact = c["ns"]
        if len(compact) >= 10:
            names.append(compact)
    names = list(dict.fromkeys(names))
    for i, row in bank.iterrows():
        if row["cr"] < 50_000:
            continue
        if row["category"] in locked:
            continue
        if row["category"] == "CUSTOMER_COLLECTION":
            continue
        compact_narr = row["ns"]
        hit = False
        for name in names:
            if name in compact_narr:
                hit = True
                break
        if hit and not inflow_interbank(compact_narr):
            bank.at[i, "category"] = "CUSTOMER_COLLECTION"


def classify_bank(bank: pd.DataFrame, customers: list[dict]) -> pd.DataFrame:
    cats = []
    buckets = []
    for row_ns, dr, cr in zip(bank["ns"], bank["dr"], bank["cr"]):
        cat, bucket = classify_row(row_ns, dr, cr)
        cats.append(cat)
        buckets.append(bucket)
    bank = bank.copy()
    bank["category"] = cats
    bank["bucket"] = buckets
    apply_name_overrides(bank, customers)
    return bank


# ---------------------------------------------------------------------------
# Matching sales / exclusions to bank credits
# ---------------------------------------------------------------------------

def match_population_to_bank(pop: pd.DataFrame, bank: pd.DataFrame, name_col: str, date_col: str, amt_col: str):
    """Greedy match of source lines to unused bank credits. Returns list of txn_id or None."""
    credits = bank[(bank["cr"] > 0) & (~bank["category"].isin(["INTERBANK", "BANK_CHARGE", "BANK_CHARGE_CONTRA"]))].copy()
    by_amt = defaultdict(list)
    for i, r in credits.iterrows():
        by_amt[round(r["cr"], 2)].append(i)

    used = set()
    match_idx = []
    for _, s in pop.iterrows():
        amt = round(float(s[amt_col]), 2)
        cands = by_amt.get(amt, [])
        st = set(tokens(s[name_col]))
        sd = s[date_col]
        best = None
        best_score = -1
        for i in cands:
            if i in used:
                continue
            bd = bank.at[i, "date"]
            nt = set(tokens(bank.at[i, "narr"]))
            ov = len(st & nt)
            days = abs((bd - sd).days) if pd.notna(sd) and pd.notna(bd) else 999
            if ov >= 2 and days <= 60:
                score = 1000 - days + ov * 5
            elif ov >= 1 and days <= 21:
                score = 600 - days
            elif days <= 3 and ov >= 1:
                score = 400 - days
            elif days <= 1 and (amt % 1_000_000) != 0 and len([c for c in cands if c not in used]) == 1:
                # Unusual amount, only one unused bank credit, same day or next day.
                score = 250 - days
            elif days == 0 and len(cands) == 1 and ov >= 0:
                # One bank credit of this amount on the same day. Accept only if the
                # sales amount is also unique that day — checked by the caller via
                # a low score that loses to any name match, and only if no other
                # candidate exists.
                score = 80
            else:
                continue
            if score > best_score:
                best_score = score
                best = i
        if best is not None:
            used.add(best)
            match_idx.append(int(bank.at[best, "txn_id"]))
        else:
            match_idx.append(None)
    return match_idx


# ---------------------------------------------------------------------------
# IFRS 15
# ---------------------------------------------------------------------------

LAND_DEFER_STATUS = {"PAYING", "PROMISE", "MIGRATED", "CONVERT TO KETU 2", "INCENTIVES", "ALMOST BOUGHT", "REFUND"}
HOUSING_RECOG_STATUS = {"BOUGHT", "COMPLETED"}


def status_u(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def project_bucket_from_name(project, product) -> str:
    p = ns(project)
    if "GORGE" in p or p == "GVC" or "GVC" in p:
        return "GVC"
    if "PACIFIC" in p:
        return "Pacific"
    if "FORESHORE" in p:
        return "Foreshore"
    if "GREEN" in p or "KETU" in p or "EPE" in p or "LANTABA" in p or "IDOBI" in p or "OMU" in p:
        return "Green City"
    if "HERITAGE" in p:
        return "Heritage"
    if "IBADAN" in p:
        return "Other land"
    if str(product).upper().startswith("HOUS"):
        return "Housing - unstated"
    return "Land - unstated"


def build_contracts(sales: pd.DataFrame, subs: pd.DataFrame) -> pd.DataFrame:
    """One row per subscription, plus uncontracted receipt groups."""
    sales = sales.copy()
    sales["cid"] = sales["Customer ID"].astype(str)
    subs = subs.copy()
    subs["cid"] = subs["Customer ID"].astype(str)
    subs["status_u"] = subs["Status"].map(status_u)
    subs["product"] = subs["Product line"].fillna("")
    subs["project"] = subs["Project / estate"].fillna("")

    assigned = set()
    rows = []

    for si, sub in subs.iterrows():
        cid = sub["cid"]
        product = str(sub["product"])
        project = str(sub["project"])
        # candidate sales: same customer, same product family
        cand_idx = []
        for i, sale in sales.iterrows():
            if i in assigned:
                continue
            if str(sale["cid"]) != cid:
                continue
            same_product = str(sale["Product line"]).upper()[:4] == product.upper()[:4]
            proj_overlap = len(set(tokens(project)) & set(tokens(sale["Project / estate"]))) >= 1
            if same_product and (proj_overlap or project == "" or pd.isna(sub["project"])):
                cand_idx.append(i)
            elif same_product and sale["Project / estate"] and project and ns(project)[:8] in ns(sale["Project / estate"]):
                cand_idx.append(i)
        # if nothing matched on project but customer has only one sub of this product, take all
        if not cand_idx:
            same_prod = [
                i for i, sale in sales.iterrows()
                if i not in assigned and str(sale["cid"]) == cid
                and str(sale["Product line"]).upper()[:4] == product.upper()[:4]
            ]
            n_subs_same = ((subs["cid"] == cid) & (subs["product"].str.upper().str[:4] == product.upper()[:4])).sum()
            if n_subs_same == 1:
                cand_idx = same_prod
        for i in cand_idx:
            assigned.add(i)
        sub_sales = sales.loc[cand_idx] if cand_idx else sales.iloc[0:0]
        cols = {y: money(sub_sales.loc[sub_sales["year"] == y, "amt"].sum()) for y in (2023, 2024, 2025)}
        rows.append({
            "kind": "subscription",
            "ref": sub["Sub ref"],
            "cid": cid,
            "name": sub["Customer name (register)"],
            "product": "Housing" if product.upper().startswith("HOUS") else "Land",
            "project": project,
            "bucket": project_bucket_from_name(project, product),
            "status": sub["status_u"],
            "sub_year": int(sub["Year"]) if pd.notna(sub["Year"]) else None,
            "contract": money(sub["Contract value (₦)"]) if pd.notna(sub["Contract value (₦)"]) else 0.0,
            "reg_paid": money(sub["Paid to date per register (₦)"]) if pd.notna(sub["Paid to date per register (₦)"]) else 0.0,
            "reg_bal": money(sub["Balance per register (₦)"]) if pd.notna(sub["Balance per register (₦)"]) else None,
            "col_2023": cols[2023], "col_2024": cols[2024], "col_2025": cols[2025],
            "allocation": sub["Allocation"] if pd.notna(sub["Allocation"]) else "",
            "docs": sub["Documents stage"] if pd.notna(sub["Documents stage"]) else "",
        })

    # uncontracted sales
    rest = sales.loc[~sales.index.isin(assigned)]
    if len(rest):
        grouped = rest.groupby(
            [rest["cid"].fillna(""), rest["Customer name (as recorded)"].fillna(""),
             rest["Product line"].fillna(""), rest["Project / estate"].fillna("")],
            dropna=False,
        )
        n = 0
        for key, g in grouped:
            n += 1
            cols = {y: money(g.loc[g["year"] == y, "amt"].sum()) for y in (2023, 2024, 2025)}
            product = "Housing" if str(key[2]).upper().startswith("HOUS") else "Land"
            rows.append({
                "kind": "uncontracted",
                "ref": f"UN-{n:04d}",
                "cid": key[0],
                "name": key[1],
                "product": product,
                "project": key[3],
                "bucket": project_bucket_from_name(key[3], product),
                "status": "",
                "sub_year": None,
                "contract": 0.0,
                "reg_paid": 0.0,
                "reg_bal": None,
                "col_2023": cols[2023], "col_2024": cols[2024], "col_2025": cols[2025],
                "allocation": "",
                "docs": "",
            })
    return pd.DataFrame(rows)


def recognise(contracts: pd.DataFrame) -> pd.DataFrame:
    """Add recognition fields and year-by-year revenue, receivable, contract liability."""
    out = contracts.copy()
    rev = {2023: [], 2024: [], 2025: []}
    rec = {2023: [], 2024: [], 2025: []}
    cl = {2023: [], 2024: [], 2025: []}
    policy = []
    recog_year = []
    revenue_total = []

    for _, c in out.iterrows():
        cols = {2023: c["col_2023"], 2024: c["col_2024"], 2025: c["col_2025"]}
        total_col = money(sum(cols.values()))
        contract = c["contract"] or 0.0
        status = c["status"] or ""
        product = c["product"]
        sub_year = c["sub_year"] if c["sub_year"] in (2022, 2023, 2024, 2025) else None

        recognise_full = False
        recognise_collected = False
        reason = ""

        if c["kind"] == "subscription" and product == "Land":
            if status in LAND_DEFER_STATUS or status == "":
                reason = "Land subscription still paying / refund / no completed status — receipts deferred"
            elif contract > 0 and (total_col >= 0.2 * contract or c["reg_paid"] >= 0.2 * contract or status in {"BOUGHT", "CLOSED", "PAID", "COMPLETED"}):
                # Do not recognise a contract value that the register shows as unpaid and unsupported
                if status == "PAID" and c["reg_paid"] == 0 and total_col < 0.2 * contract:
                    reason = "Land marked paid but no collections evidenced — deferred"
                else:
                    recognise_full = True
                    reason = "Land control transferred (allocated / bought). Revenue at contract value"
            else:
                recognise_collected = True
                reason = "Land sale evidenced by receipts; contract value not reliable — revenue equals collections"
        elif c["kind"] == "subscription" and product == "Housing":
            paid_ratio = (max(total_col, c["reg_paid"]) / contract) if contract else 0
            if status in HOUSING_RECOG_STATUS and contract > 0 and paid_ratio >= 0.5:
                recognise_full = True
                reason = "Housing handover evidenced (Bought/Completed) and substantially paid — revenue at contract value"
            elif status == "PAID" and contract > 0 and c["reg_paid"] >= 0.8 * contract:
                recognise_full = True
                reason = "Housing marked paid and register paid-to-date supports handover — revenue at contract value"
            else:
                reason = "Housing not yet handed over — receipts held as deferred income (contract liability)"
        else:
            # uncontracted
            if product == "Land" and total_col > 0:
                recognise_collected = True
                reason = "Plot receipt with no open subscription — recognised in the year of receipt (turnover-register basis)"
            elif product == "Housing":
                reason = "Housing receipt with no handover evidence — deferred income"
            else:
                reason = "Deferred"

        if recognise_full:
            rev_amount = contract
            # recognition year: land — subscription year (or first collection year if sub year missing / 2022)
            if product == "Land":
                ry = sub_year if sub_year in (2023, 2024, 2025) else None
                if ry is None:
                    ry = next((y for y in (2023, 2024, 2025) if cols[y] > 0), 2023)
            else:
                # housing: year cumulative collections (or register) reach 50% of contract, else last collection year
                running = 0.0
                ry = None
                for y in (2023, 2024, 2025):
                    running += cols[y]
                    if contract and running >= 0.5 * contract:
                        ry = y
                        break
                if ry is None:
                    ry = next((y for y in (2025, 2024, 2023) if cols[y] > 0), sub_year or 2025)
                    if ry not in (2023, 2024, 2025):
                        ry = 2025
        elif recognise_collected:
            rev_amount = total_col
            ry = None  # recognised as collected, by year
        else:
            rev_amount = 0.0
            ry = None

        policy.append(reason)
        recog_year.append(ry if ry else "")
        revenue_total.append(rev_amount)

        cum = 0.0
        for y in (2023, 2024, 2025):
            cum = money(cum + cols[y])
            if recognise_collected:
                rev_y = cols[y]
                rec_y = 0.0
                cl_y = 0.0
            elif recognise_full:
                if y < ry:
                    rev_y = 0.0
                    rec_y = 0.0
                    cl_y = cum
                else:
                    rev_y = rev_amount if y == ry else 0.0
                    rec_y = money(max(0.0, rev_amount - cum))
                    cl_y = money(max(0.0, cum - rev_amount))
                    # cap receivable at register balance when the register is populated and lower
                    if c["reg_bal"] is not None and c["reg_paid"] > 0 and c["reg_bal"] >= 0:
                        rec_y = money(min(rec_y, c["reg_bal"]))
                        # if we cap receivable, the difference stays as unbilled? 
                        # revenue was contract; if we reduce receivable we must reduce revenue
                        # to keep the identity. Adjust revenue in the recognition year.
                        if y == ry:
                            # identity: rev = cum_at_ry + rec + cl_future_wait
                            # At recognition, cl should be max(0, cum - rev_amount) which is 0 if rev=contract>=cum
                            # If we cap rec below contract-cum, reduce rev_amount for the identity.
                            shortfall = money(max(0.0, (rev_amount - cum) - rec_y))
                            if shortfall > 0:
                                rev_y = money(rev_y - shortfall)
                                rev_amount = money(rev_amount - shortfall)
            else:
                rev_y = 0.0
                rec_y = 0.0
                cl_y = cum
            rev[y].append(money(rev_y))
            rec[y].append(money(rec_y))
            cl[y].append(money(cl_y))

    out["policy"] = policy
    out["recog_year"] = recog_year
    out["revenue_total"] = revenue_total
    for y in (2023, 2024, 2025):
        out[f"rev_{y}"] = rev[y]
        out[f"rec_{y}"] = rec[y]
        out[f"cl_{y}"] = cl[y]
    return out


# ---------------------------------------------------------------------------
# Ledger
# ---------------------------------------------------------------------------

class Ledger:
    def __init__(self):
        self.rows = []

    def post(self, year, dr, cr, amount, memo, source=""):
        amount = money(amount)
        if amount == 0:
            return
        if amount < 0:
            dr, cr = cr, dr
            amount = -amount
        self.rows.append({
            "year": int(year), "dr": dr, "cr": cr, "amount": amount,
            "memo": memo, "source": source,
        })

    def activity(self, account, year=None, upto=False, exclude_source=None):
        dr = cr = 0.0
        for r in self.rows:
            if exclude_source and r["source"] in exclude_source:
                continue
            if year is not None:
                if upto and r["year"] > year:
                    continue
                if not upto and r["year"] != year:
                    continue
            if r["dr"] == account:
                dr += r["amount"]
            if r["cr"] == account:
                cr += r["amount"]
        return money(dr), money(cr)

    def accounts(self):
        names = set()
        for r in self.rows:
            names.add(r["dr"])
            names.add(r["cr"])
        return sorted(names)

    def balance(self, account, year, normal="dr"):
        dr, cr = self.activity(account, year, upto=True)
        net = money(dr - cr)
        return net if normal == "dr" else money(cr - dr)


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def build():
    print("Loading bank...")
    bank = load_bank()
    print(f"  bank rows {len(bank):,}  cr {bank['cr'].sum():,.2f}  dr {bank['dr'].sum():,.2f}")
    print("Loading sales pack...")
    sales, subs, cust, excl, exceptions = load_sales_pack()
    print(f"  sales {len(sales):,} {sales['amt'].sum():,.2f}  subs {len(subs):,}  excl {len(excl):,}")

    customers = []
    for col, frame in (
        ("Customer name (as recorded)", sales),
        ("Customer name (register)", subs),
        ("Customer name", cust),
    ):
        for name in frame[col].dropna().unique():
            customers.append({"name": str(name), "ns": ns(name), "tokens": tokens(name)})
    # realtors
    for name in subs["Realtor / contact"].dropna().unique():
        customers.append({"name": str(name), "ns": ns(name), "tokens": tokens(name)})

    print("Classifying bank...")
    bank = classify_bank(bank, customers)

    print("Matching sales ledger to bank credits...")
    sales = sales.reset_index(drop=True)
    sales["bank_txn"] = match_population_to_bank(sales, bank, "Customer name (as recorded)", "date", "amt")
    traced = sales["bank_txn"].notna().sum()
    traced_amt = sales.loc[sales["bank_txn"].notna(), "amt"].sum()
    print(f"  traced {traced}/{len(sales)} lines  N{traced_amt:,.0f}  untraced N{sales['amt'].sum()-traced_amt:,.0f}")

    # Force matched credits into customer collections (they are property receipts)
    matched_ids = set(int(x) for x in sales["bank_txn"].dropna().tolist())
    bank.loc[bank["txn_id"].isin(matched_ids), "category"] = "CUSTOMER_COLLECTION"
    bank.loc[bank["txn_id"].isin(matched_ids), "matched_sale"] = True

    print("Matching exclusions...")
    excl = excl.reset_index(drop=True)
    excl["bank_txn"] = match_population_to_bank(
        excl, bank, "Customer", "date", "amt"
    )
    # do not steal a credit already matched to a sale
    excl.loc[excl["bank_txn"].isin(matched_ids), "bank_txn"] = None
    excl_ids = set(int(x) for x in excl["bank_txn"].dropna().tolist())
    for i, r in excl.iterrows():
        if pd.isna(r["bank_txn"]):
            continue
        reason = str(r.get("Reason excluded", ""))
        cat = "CO_OWNERSHIP_IN"
        if "Service" in reason or "service" in reason:
            cat = "OTHER_INCOME_SERVICE"
        if "Credit memo" in reason or "non-cash" in reason:
            continue
        bank.loc[bank["txn_id"] == int(r["bank_txn"]), "category"] = cat

    print("Building IFRS 15 contracts...")
    contracts = build_contracts(sales, subs)
    contracts = recognise(contracts)

    # Additional bank collections: customer-collection credits not matched to a sale
    # and not an exclusion match. These increase deferred income (not revenue),
    # unless a later manual review ties them to a contract. Kept separate so the
    # sales-ledger turnover is not mixed with unidentified lookalikes.
    addl_mask = (
        (bank["category"] == "CUSTOMER_COLLECTION")
        & (~bank["txn_id"].isin(matched_ids))
        & (~bank["txn_id"].isin(excl_ids))
        & (bank["cr"] > 0)
    )
    addl = bank.loc[addl_mask]
    addl_by_year = {y: money(addl.loc[addl["year"] == y, "cr"].sum()) for y in (2023, 2024, 2025)}
    print("  additional bank property receipts not in sales ledger", addl_by_year)

    # ---- ledger ----
    L = Ledger()
    # Opening
    L.post(2023, "Cash", "Share capital", SHARE_CAPITAL, "Paid-up ordinary share capital brought forward", "OPEN")
    L.post(2023, "Cash", "Retained earnings", money(OPENING_CASH_TOTAL - SHARE_CAPITAL),
           "Opening net assets limited to verified bank cash, after share capital", "OPEN")
    # The second post may be negative (RE debit). Ledger.post flips negative amounts.
    # OPENING_CASH_TOTAL 478,080 - 1,000,000 = -521,920, so post() will Dr RE Cr Cash.
    # That would REDUCE cash. Wrong. Post opening cash explicitly.
    # Remove the implicit effect by posting cash as a single opening entry.
    # Easier: clear and redo opening correctly.
    L.rows = [r for r in L.rows if r["source"] != "OPEN"]
    # Cash 478,080 = share capital 1,000,000 + retained earnings (521,920).
    L.post(2023, "Cash", "Opening equity clearing", OPENING_CASH_TOTAL, "Bank cash at 1 January 2023", "OPEN")
    L.post(2023, "Opening equity clearing", "Share capital", SHARE_CAPITAL, "Paid-up share capital", "OPEN")
    L.post(2023, "Retained earnings", "Opening equity clearing", money(SHARE_CAPITAL - OPENING_CASH_TOTAL),
           "Opening retained earnings (deficit) — cash less share capital", "OPEN")

    # Bank movements
    for _, r in bank.iterrows():
        year = int(r["year"])
        cat = r["category"]
        if r["cr"] > 0:
            L.post(year, "Cash", f"CAT::{cat}", r["cr"], r["narr"][:180], "BANK")
        if r["dr"] > 0:
            L.post(year, f"CAT::{cat}", "Cash", r["dr"], r["narr"][:180], "BANK")

    # Align the customer-collection account to the sales ledger on a cumulative basis.
    # Line-by-line misses are not booked twice (once as untraced and once as extra
    # bank receipts). Only the cumulative gap is recognised.
    bank_cc = {
        y: money(bank.loc[(bank["year"] == y) & (bank["category"] == "CUSTOMER_COLLECTION"), "cr"].sum())
        for y in (2023, 2024, 2025)
    }
    sales_by_year = {y: money(sales.loc[sales["year"] == y, "amt"].sum()) for y in (2023, 2024, 2025)}
    cum_s = cum_b = 0.0
    prev_u = prev_a = 0.0
    gap_untraced = {}
    gap_addl = {}
    for y in (2023, 2024, 2025):
        cum_s = money(cum_s + sales_by_year[y])
        cum_b = money(cum_b + bank_cc[y])
        u_close = money(max(0.0, cum_s - cum_b))
        a_close = money(max(0.0, cum_b - cum_s))
        du = money(u_close - prev_u)
        da = money(a_close - prev_a)
        if du > 0:
            L.post(y, "Untraced customer receipts", "CAT::CUSTOMER_COLLECTION", du,
                   "Sales ledger exceeds property receipts identified in the four banks", "GAP")
        elif du < 0:
            L.post(y, "CAT::CUSTOMER_COLLECTION", "Untraced customer receipts", -du,
                   "Bank receipts now cover sales previously untraced", "GAP")
        if da > 0:
            L.post(y, "CAT::CUSTOMER_COLLECTION", "Contract liabilities", da,
                   "Bank property receipts in excess of the sales ledger — deferred", "GAP")
        elif da < 0:
            L.post(y, "Contract liabilities", "CAT::CUSTOMER_COLLECTION", -da,
                   "Release of excess bank receipts previously deferred", "GAP")
        prev_u, prev_a = u_close, a_close
        gap_untraced[y] = u_close
        gap_addl[y] = a_close
    addl_by_year = gap_addl
    print("  cumulative gap untraced", gap_untraced, "excess bank deferred", gap_addl)

    # Untraced co-ownership (exclusions not matched, not credit memos)
    for _, r in excl.iterrows():
        reason = str(r.get("Reason excluded", ""))
        if "Credit memo" in reason or "non-cash" in reason:
            continue
        if pd.notna(r["bank_txn"]):
            continue
        year = int(r["year"]) if pd.notna(r["year"]) else 2024
        if "Service" in reason or "service" in reason:
            L.post(year, "Untraced other receipts", "CAT::OTHER_INCOME_SERVICE", r["amt"],
                   "Service income in source register, not traced to bank", "UNTRACED_EXCL")
        else:
            L.post(year, "Untraced other receipts", "CAT::CO_OWNERSHIP_IN", r["amt"],
                   "Co-ownership receipt in source register, not traced to bank", "UNTRACED_EXCL")

    # IFRS 15 entries from contracts, by year
    for y in (2023, 2024, 2025):
        rev_land = money(contracts.loc[contracts["product"] == "Land", f"rev_{y}"].sum())
        rev_hous = money(contracts.loc[contracts["product"] == "Housing", f"rev_{y}"].sum())
        # movement in receivable and CL
        rec_close = money(contracts[f"rec_{y}"].sum())
        cl_close = money(contracts[f"cl_{y}"].sum())
        rec_open = money(contracts[f"rec_{y-1}"].sum()) if y > 2023 else 0.0
        cl_open = money(contracts[f"cl_{y-1}"].sum()) if y > 2023 else 0.0
        d_rec = money(rec_close - rec_open)
        d_cl = money(cl_close - cl_open)
        # Identity: revenue = collections_this_year + d_rec - d_cl
        # We credit revenue and CL movement against the customer-collection account
        # and debit receivable movement.
        L.post(y, "CAT::CUSTOMER_COLLECTION", "Revenue - land", rev_land, "IFRS 15 land revenue", "IFRS")
        L.post(y, "CAT::CUSTOMER_COLLECTION", "Revenue - housing", rev_hous, "IFRS 15 housing revenue", "IFRS")
        if d_rec >= 0:
            L.post(y, "Trade receivables", "CAT::CUSTOMER_COLLECTION", d_rec, "Movement in earned unpaid balances", "IFRS")
        else:
            L.post(y, "CAT::CUSTOMER_COLLECTION", "Trade receivables", -d_rec, "Collections of prior receivables", "IFRS")
        if d_cl >= 0:
            L.post(y, "CAT::CUSTOMER_COLLECTION", "Contract liabilities", d_cl, "Increase in deferred income", "IFRS")
        else:
            L.post(y, "Contract liabilities", "CAT::CUSTOMER_COLLECTION", -d_cl, "Release of deferred income", "IFRS")

        # Reclass remaining category accounts into FS lines.
    # Customer collection should now be ~0. Any residual is a bridge difference.
    # Map categories:
    mapping_cr = {
        # credit-balance categories (liabilities / income) already posted as Cr CAT
    }
    # We leave CAT::* in the ledger and map them at presentation, except we
    # reclass a few into named accounts for a cleaner TB.
    reclass = {
        "CO_OWNERSHIP_IN": ("Contract liabilities - co-ownership", "cr"),
        "CO_OWNERSHIP_OUT": ("Contract liabilities - co-ownership", "dr"),
        "RELATED_PARTY_IN": ("Related party current account", "cr"),
        "RELATED_PARTY_OUT": ("Related party current account", "dr"),
        "BORROWINGS_IN": ("Borrowings", "cr"),
        "BORROWINGS_OUT": ("Borrowings", "dr"),
        "UNALLOCATED_IN": ("Receipts pending allocation", "cr"),
        "UNALLOCATED_OUT": ("Payments pending allocation", "dr"),
        "FX_ADVANCE": ("FX and other advances", "dr"),
        "TAX_PAID": ("Current tax", "dr"),
        "INTERBANK": ("Inter-account transfers", "net"),
        "PPE": ("PPE cost", "dr"),
        "INVENTORY_LAND": ("Inventory", "dr"),
        "INVENTORY_HOUSING": ("Inventory", "dr"),
        "INVENTORY_DEVELOPMENT": ("Inventory", "dr"),
        "INVENTORY_CONTRA": ("Inventory", "cr"),
        "DIVIDEND": ("Dividends", "dr"),
        "TRADING_PURCHASES": ("Cost of sales - trading", "dr"),
        "OTHER_INCOME_TRADING": ("Other income - trading", "cr"),
        "OTHER_INCOME_SERVICE": ("Other income - service", "cr"),
        "OTHER_INCOME_REFUND": ("Other income - refunds", "cr"),
        "CUSTOMER_REFUND": ("Customer refunds", "dr"),
        "STAFF_COST": ("Staff costs", "dr"),
        "DIRECTOR_REMUNERATION": ("Director remuneration", "dr"),
        "SELLING_COMMISSION": ("Selling commissions", "dr"),
        "RENT": ("Rent", "dr"),
        "PROFESSIONAL": ("Professional fees", "dr"),
        "ADMIN_OPS": ("Administrative expenses", "dr"),
        "BANK_CHARGE": ("Bank charges", "dr"),
        "BANK_CHARGE_CONTRA": ("Bank charges", "cr"),
        "FINANCE_COST": ("Finance costs", "dr"),
        "FINANCE_INCOME": ("Finance income", "cr"),
    }

    # Snapshot category net by year then post reclass and drop CAT from presentation
    # by posting the reverse of remaining CAT balances into the mapped accounts.
    for y in (2023, 2024, 2025):
        cats = sorted({r["dr"][5:] for r in L.rows if r["dr"].startswith("CAT::") and r["year"] == y}
                      | {r["cr"][5:] for r in L.rows if r["cr"].startswith("CAT::") and r["year"] == y})
        for cat in cats:
            if cat not in reclass:
                continue
            acct, side = reclass[cat]
            dr, cr = L.activity(f"CAT::{cat}", y, upto=False)
            # net debit in the category account
            net_dr = money(dr - cr)
            if abs(net_dr) < 0.01:
                continue
            if side == "dr":
                # category should be a debit. Move net_dr from CAT to acct.
                if net_dr > 0:
                    L.post(y, acct, f"CAT::{cat}", net_dr, f"Reclass {cat}", "RECLASS")
                else:
                    L.post(y, f"CAT::{cat}", acct, -net_dr, f"Reclass {cat}", "RECLASS")
            elif side == "cr":
                net_cr = money(cr - dr)
                if net_cr > 0:
                    L.post(y, f"CAT::{cat}", acct, net_cr, f"Reclass {cat}", "RECLASS")
                else:
                    L.post(y, acct, f"CAT::{cat}", -net_cr, f"Reclass {cat}", "RECLASS")
            elif side == "net":
                # interbank: net debit = asset (funds sent to accounts not mirrored), net credit = liability
                if net_dr > 0:
                    L.post(y, acct, f"CAT::{cat}", net_dr, "Net inter-account outflow", "RECLASS")
                elif net_dr < 0:
                    L.post(y, f"CAT::{cat}", acct, -net_dr, "Net inter-account inflow", "RECLASS")

    # Inventory release to cost of sales — proportional to recognised revenue
    # versus revenue + contract liability, by project bucket.
    print("Releasing inventory to cost of sales...")
    # inventory additions by year and bucket from bank
    inv_add = {y: defaultdict(float) for y in (2023, 2024, 2025)}
    for _, r in bank.iterrows():
        if r["category"] in ("INVENTORY_LAND", "INVENTORY_HOUSING", "INVENTORY_DEVELOPMENT") and r["dr"] > 0:
            bucket = r["bucket"] or "General development"
            inv_add[int(r["year"])][bucket] += r["dr"]

    # revenue and CL by bucket from contracts (CL here excludes additional bank receipts,
    # which are not project-tagged; they stay out of the allocation base)
    def bucket_rev_cl(year):
        rev_b = defaultdict(float)
        cl_b = defaultdict(float)
        for _, c in contracts.iterrows():
            rev_b[c["bucket"]] += c[f"rev_{year}"]
            cl_b[c["bucket"]] += c[f"cl_{year}"]
        return rev_b, cl_b

    released_cum = defaultdict(float)
    cos_by_year = {}
    inventory_by_bucket_close = {}
    for y in (2023, 2024, 2025):
        rev_b, cl_b = bucket_rev_cl(y)
        # cumulative cost
        cum_cost = defaultdict(float)
        for yy in (2023, 2024, 2025):
            if yy > y:
                break
            for bkt, amt in inv_add[yy].items():
                cum_cost[bkt] += amt
        cos_y = defaultdict(float)
        close_bal = {}
        buckets = set(cum_cost) | set(rev_b) | set(cl_b)
        for bkt in buckets:
            cost = cum_cost[bkt]
            earned = rev_b.get(bkt, 0.0)
            unearned = cl_b.get(bkt, 0.0)
            base = earned + unearned
            if cost <= 0:
                target = 0.0
            elif base <= 0:
                target = 0.0  # cost with no customer activity remains inventory
            else:
                target = cost * earned / base
            target = min(cost, target)
            release = money(target - released_cum[bkt])
            if release < 0:
                release = 0.0  # do not reverse prior releases on a change in mix
            released_cum[bkt] = money(released_cum[bkt] + release)
            cos_y[bkt] = release
            close_bal[bkt] = money(cost - released_cum[bkt])
        cos_by_year[y] = {k: money(v) for k, v in cos_y.items()}
        inventory_by_bucket_close[y] = close_bal
        total_cos = money(sum(cos_y.values()))
        L.post(y, "Cost of sales", "Inventory", total_cos, "Project cost released in proportion to revenue recognised", "COS")

    # Depreciation — 20% straight line on PPE additions, full year in year of purchase.
    ppe_add = {y: money(bank.loc[(bank["year"] == y) & (bank["category"] == "PPE"), "dr"].sum()) for y in (2023, 2024, 2025)}
    accum = 0.0
    dep_by_year = {}
    for y in (2023, 2024, 2025):
        accum = money(accum + ppe_add[y])
        dep = money(accum * 0.20)
        # charge only the increment
        prior = dep_by_year.get(y - 1, 0.0)
        # prior stored as charge not balance. Recompute charge as 20% of gross cost to date,
        # minus depreciation already charged.
        charge = dep  # this is the full required accum; charge = required accum - prior accum
        # fix below after loop structure
        dep_by_year[y] = dep
    charged = 0.0
    for y in (2023, 2024, 2025):
        charge = money(dep_by_year[y] - charged)
        charged = dep_by_year[y]
        dep_by_year[y] = charge
        L.post(y, "Depreciation", "Accumulated depreciation", charge, "20% straight line on identified PPE", "DEP")

    # ECL — 5% of gross trade receivables, movement through P&L
    ecl_rate = 0.05
    prior_allow = 0.0
    ecl_by_year = {}
    for y in (2023, 2024, 2025):
        gross_rec = money(contracts[f"rec_{y}"].sum())
        allow = money(gross_rec * ecl_rate)
        movement = money(allow - prior_allow)
        ecl_by_year[y] = {"gross": gross_rec, "allowance": allow, "movement": movement}
        if movement >= 0:
            L.post(y, "Impairment of receivables", "ECL allowance", movement, "General ECL 5% — no loss history in the source files", "ECL")
        else:
            L.post(y, "ECL allowance", "Impairment of receivables", -movement, "Reduction in ECL allowance", "ECL")
        prior_allow = allow

    # Tax
    print("Computing tax...")
    tax = {}
    prior_tax_liab = 0.0
    for y in (2023, 2024, 2025):
        def pnl(acct, year=y):
            d, c = L.activity(acct, year, upto=False)
            return money(c - d)  # income positive

        def exp(acct, year=y):
            d, c = L.activity(acct, year, upto=False)
            return money(d - c)

        rev_land = pnl("Revenue - land")
        rev_hous = pnl("Revenue - housing")
        oth_tr = pnl("Other income - trading")
        oth_sv = pnl("Other income - service")
        oth_rf = pnl("Other income - refunds")
        fin_in = pnl("Finance income")
        revenue = money(rev_land + rev_hous)
        other_income = money(oth_tr + oth_sv + oth_rf)
        turnover = money(revenue + other_income)
        cos = exp("Cost of sales") + exp("Cost of sales - trading")
        refunds = exp("Customer refunds")
        opex = money(
            exp("Staff costs") + exp("Director remuneration") + exp("Selling commissions")
            + exp("Rent") + exp("Professional fees") + exp("Administrative expenses")
            + exp("Bank charges") + exp("Depreciation") + exp("Impairment of receivables")
            + refunds
        )
        fin_cost = exp("Finance costs")
        pbt = money(revenue + other_income + fin_in - cos - opex - fin_cost)

        if turnover <= 25_000_000:
            rate, size, small = 0.0, "Small", True
        elif turnover <= 100_000_000:
            rate, size, small = 0.20, "Medium", False
        else:
            rate, size, small = 0.30, "Large", False

        dep = exp("Depreciation")
        ecl_mv = exp("Impairment of receivables")
        assessable = money(pbt + dep + max(ecl_mv, 0))
        taxable = max(assessable, 0.0)
        cit = money(taxable * rate)
        min_tax = 0.0 if small else money(0.005 * max(turnover, 0))
        min_applied = (not small) and cit < min_tax
        cit_payable = min_tax if min_applied else cit
        tet = 0.0 if small or assessable <= 0 else money(0.03 * assessable)
        ptf = 0.0 if pbt <= 0 else money(0.00005 * pbt)
        naseni = 0.0 if small or turnover < 100_000_000 or pbt <= 0 else money(0.0025 * pbt)
        current = money(cit_payable + tet + ptf + naseni)
        L.post(y, "Income tax expense", "Current tax", current, f"CIT/TET/PTF/NASENI {size}", "TAX")

        tax_paid = exp_upto_movement("Current tax", L, y)  # placeholder replaced below
        tax[y] = {
            "revenue": revenue, "other_income": other_income, "turnover": turnover,
            "cos": cos, "opex": opex, "fin_cost": fin_cost, "fin_in": fin_in,
            "pbt": pbt, "size": size, "rate": rate, "assessable": assessable,
            "cit": cit, "min_tax": min_tax, "min_applied": min_applied,
            "cit_payable": cit_payable, "tet": tet, "ptf": ptf, "naseni": naseni,
            "current": current, "dep": dep, "ecl": ecl_mv, "refunds": refunds,
            "rev_land": rev_land, "rev_hous": rev_hous,
        }

    # Close P&L to retained earnings
    pnl_accounts = [
        "Revenue - land", "Revenue - housing", "Other income - trading",
        "Other income - service", "Other income - refunds", "Finance income",
        "Cost of sales", "Cost of sales - trading", "Customer refunds",
        "Staff costs", "Director remuneration", "Selling commissions", "Rent",
        "Professional fees", "Administrative expenses", "Bank charges",
        "Depreciation", "Impairment of receivables", "Finance costs",
        "Income tax expense",
    ]
    for y in (2023, 2024, 2025):
        for acct in pnl_accounts:
            d, c = L.activity(acct, y, upto=False)
            net_cr = money(c - d)
            if abs(net_cr) < 0.01:
                continue
            if net_cr > 0:
                L.post(y, acct, "Retained earnings", net_cr, "Close to retained earnings", "CLOSE")
            else:
                L.post(y, "Retained earnings", acct, -net_cr, "Close to retained earnings", "CLOSE")
        # Dividends are an appropriation, not an expense.
        d, c = L.activity("Dividends", y, upto=False)
        net_dr = money(d - c)
        if abs(net_dr) >= 0.01:
            L.post(y, "Retained earnings", "Dividends", net_dr, "Dividend appropriation", "CLOSE")

    # Cash proof against stated balances
    stated = stated_cash_balances(bank)
    cash_proof = {}
    rolled = OPENING_CASH_TOTAL
    for y in (2023, 2024, 2025):
        d, c = L.activity("Cash", y, upto=False)
        # opening entry is in 2023 and includes opening cash as a debit
        rolled = money(L.balance("Cash", y, "dr"))
        cash_proof[y] = {
            "ledger": rolled,
            "stated": stated[y]["total"],
            "diff": money(stated[y]["total"] - rolled),
            "by_bank": stated[y]["by_bank"],
        }
    # Post the difference so the statement cash equals the bank. A non-zero
    # difference is disclosed; it is not revenue.
    for y in (2023, 2024, 2025):
        # movement difference vs prior stated
        prior = stated[y - 1]["total"] if y > 2023 else OPENING_CASH_TOTAL
        # ledger cash already includes opening. Compare closing.
        diff = cash_proof[y]["diff"]
        prev_diff = cash_proof[y - 1]["diff"] if y > 2023 else 0.0
        incr = money(diff - prev_diff)
        if abs(incr) >= 0.01:
            if incr > 0:
                L.post(y, "Cash", "Bank statement difference", incr, "Align cash to stated bank balance", "CASHPROOF")
            else:
                L.post(y, "Bank statement difference", "Cash", -incr, "Align cash to stated bank balance", "CASHPROOF")
            # close the P&L-neutral difference: it is a BS account. If we credited
            # income-like "Bank statement difference", keep it as a BS liability/asset.
        cash_proof[y]["diff_posted_cumulative"] = diff

    # Re-close nothing: bank statement difference is BS.

    model = assemble(
        L, bank, sales, subs, cust, excl, exceptions, contracts,
        addl_by_year, cash_proof, stated, tax, ecl_by_year, ppe_add,
        dep_by_year, cos_by_year, inventory_by_bucket_close, inv_add,
        traced_amt, matched_ids,
    )
    return model


def exp_upto_movement(account, L, year):
    return 0.0


def stated_cash_balances(bank: pd.DataFrame) -> dict:
    out = {2022: {"total": OPENING_CASH_TOTAL, "by_bank": dict(OPENING_CASH)}}
    for y in (2023, 2024, 2025):
        by = {}
        for bname, opening in OPENING_CASH.items():
            sub = bank[(bank["bank"] == bname) & (bank["date"] <= pd.Timestamp(f"{y}-12-31"))]
            if len(sub) == 0:
                by[bname] = opening if y == 2023 or bname != "GTBank" else 0.0
                if bname == "GTBank" and y == 2023:
                    by[bname] = 0.0
                elif len(sub) == 0:
                    by[bname] = opening
            else:
                last = sub.iloc[-1]
                # files are not guaranteed sorted within a day; take last date's last row
                last_day = sub[sub["date"] == sub["date"].max()]
                by[bname] = money(last_day.iloc[-1]["bal"])
        # GTBank did not exist in 2023
        if y == 2023:
            by["GTBank"] = 0.0
        out[y] = {"by_bank": by, "total": money(sum(by.values()))}
    return out


def assemble(L, bank, sales, subs, cust, excl, exceptions, contracts,
             addl_by_year, cash_proof, stated, tax, ecl_by_year, ppe_add,
             dep_by_year, cos_by_year, inv_close, inv_add, traced_amt, matched_ids):
    years = (2023, 2024, 2025)

    def bal(acct, year, normal="dr"):
        return L.balance(acct, year, normal)

    def yr_amt(acct, year, normal="dr"):
        d, c = L.activity(acct, year, upto=False, exclude_source={"CLOSE"})
        return money(d - c) if normal == "dr" else money(c - d)

    sfp = {}
    pnl = {}
    for y in years:
        cash = bal("Cash", y, "dr")
        rec_g = bal("Trade receivables", y, "dr")
        ecl = bal("ECL allowance", y, "cr")
        untraced = bal("Untraced customer receipts", y, "dr")
        untraced_other = bal("Untraced other receipts", y, "dr")
        inventory = bal("Inventory", y, "dr")
        ppe = bal("PPE cost", y, "dr")
        accdep = bal("Accumulated depreciation", y, "cr")
        pending_pay = bal("Payments pending allocation", y, "dr")
        fx = bal("FX and other advances", y, "dr")
        bank_diff = bal("Bank statement difference", y, "cr")  # credit normal if we credited it
        # bank diff may be debit or credit
        bank_diff_dr = bal("Bank statement difference", y, "dr")
        inter = bal("Inter-account transfers", y, "dr")
        rp = bal("Related party current account", y, "dr")  # debit = due from director
        rp_cr = bal("Related party current account", y, "cr")
        borrow_cr = bal("Borrowings", y, "cr")
        borrow_dr = bal("Borrowings", y, "dr")
        cl = bal("Contract liabilities", y, "cr")
        co = bal("Contract liabilities - co-ownership", y, "cr")
        pending_rec = bal("Receipts pending allocation", y, "cr")
        tax_cr = bal("Current tax", y, "cr")
        tax_dr = bal("Current tax", y, "dr")

        sfp[y] = {
            "cash": cash,
            "receivables_gross": rec_g,
            "ecl": ecl,
            "receivables_net": money(rec_g - ecl),
            "untraced_receipts": untraced,
            "untraced_other": untraced_other,
            "inventory": inventory,
            "ppe_cost": ppe,
            "accum_dep": accdep,
            "ppe_nbv": money(ppe - accdep),
            "payments_pending": pending_pay,
            "fx_advances": fx,
            "interaccount_dr": inter if inter > 0 else 0.0,
            "related_party_dr": rp if rp > 0 else 0.0,
            "borrowings_dr": borrow_dr if borrow_dr > 0 else 0.0,
            "bank_diff_dr": bank_diff_dr if bank_diff_dr > 0 else 0.0,
            "tax_dr": tax_dr if tax_dr > 0 else 0.0,
            "contract_liabilities": cl,
            "co_ownership": co if co > 0 else 0.0,
            "co_ownership_dr": -co if co < 0 else 0.0,
            "receipts_pending": pending_rec if pending_rec > 0 else 0.0,
            "receipts_pending_dr": -pending_rec if pending_rec < 0 else 0.0,
            "related_party_cr": rp_cr if rp_cr > 0 else 0.0,
            "borrowings_cr": borrow_cr if borrow_cr > 0 else 0.0,
            "tax_cr": tax_cr if tax_cr > 0 else 0.0,
            "bank_diff_cr": bank_diff if bank_diff > 0 else 0.0,
            "interaccount_cr": -inter if inter < 0 else 0.0,
            "share_capital": bal("Share capital", y, "cr"),
            "retained_earnings": bal("Retained earnings", y, "cr"),
        }
        # assets and liabilities assembled in presentation with sign-safe fields
        pnl[y] = {
            "rev_land": yr_amt("Revenue - land", y, "cr"),
            "rev_hous": yr_amt("Revenue - housing", y, "cr"),
            "other_trading": yr_amt("Other income - trading", y, "cr"),
            "other_service": yr_amt("Other income - service", y, "cr"),
            "other_refunds": yr_amt("Other income - refunds", y, "cr"),
            "finance_income": yr_amt("Finance income", y, "cr"),
            "cos": yr_amt("Cost of sales", y, "dr") + yr_amt("Cost of sales - trading", y, "dr"),
            "cos_trading": yr_amt("Cost of sales - trading", y, "dr"),
            "refunds": yr_amt("Customer refunds", y, "dr"),
            "staff": yr_amt("Staff costs", y, "dr"),
            "director_rem": yr_amt("Director remuneration", y, "dr"),
            "commission": yr_amt("Selling commissions", y, "dr"),
            "rent": yr_amt("Rent", y, "dr"),
            "professional": yr_amt("Professional fees", y, "dr"),
            "admin": yr_amt("Administrative expenses", y, "dr"),
            "bank_charges": yr_amt("Bank charges", y, "dr"),
            "depreciation": yr_amt("Depreciation", y, "dr"),
            "ecl_expense": yr_amt("Impairment of receivables", y, "dr"),
            "finance_cost": yr_amt("Finance costs", y, "dr"),
            "tax": yr_amt("Income tax expense", y, "dr"),
        }
        p = pnl[y]
        p["revenue"] = money(p["rev_land"] + p["rev_hous"])
        p["other_income"] = money(p["other_trading"] + p["other_service"] + p["other_refunds"])
        p["gross_profit"] = money(p["revenue"] + p["other_income"] - p["cos"] - p["refunds"])
        p["opex"] = money(p["staff"] + p["director_rem"] + p["commission"] + p["rent"]
                          + p["professional"] + p["admin"] + p["bank_charges"]
                          + p["depreciation"] + p["ecl_expense"])
        p["pbt"] = money(p["gross_profit"] - p["opex"] + p["finance_income"] - p["finance_cost"])
        p["pat"] = money(p["pbt"] - p["tax"])
        p["dividends"] = yr_amt("Dividends", y, "dr")

    # Category totals
    cat = (
        bank.groupby(["year", "category"])[["dr", "cr"]]
        .sum()
        .reset_index()
    )

    # Collections bridge
    collections = {y: money(sales.loc[sales["year"] == y, "amt"].sum()) for y in years}
    untraced = {
        y: money(sales.loc[(sales["year"] == y) & sales["bank_txn"].isna(), "amt"].sum())
        for y in years
    }
    cl_total = {y: money(contracts[f"cl_{y}"].sum() + addl_by_year[y]) for y in years}
    rec_total = {y: money(contracts[f"rec_{y}"].sum()) for y in years}

    # Top unallocated
    un_in = bank[(bank["category"] == "UNALLOCATED_IN") & (bank["cr"] >= 5_000_000)][
        ["date", "bank", "cr", "narr", "year"]
    ].sort_values("cr", ascending=False)
    un_out = bank[(bank["category"] == "UNALLOCATED_OUT") & (bank["dr"] >= 5_000_000)][
        ["date", "bank", "dr", "narr", "year"]
    ].sort_values("dr", ascending=False)

    # Balance check
    checks = []
    for y in years:
        s = sfp[y]
        assets = money(
            s["cash"] + s["receivables_net"] + s["untraced_receipts"] + s["untraced_other"] + s["inventory"]
            + s["ppe_nbv"] + s["payments_pending"] + s["fx_advances"]
            + s["interaccount_dr"] + s["related_party_dr"] + s["borrowings_dr"]
            + s["bank_diff_dr"] + s["tax_dr"] + s["co_ownership_dr"]
            + s["receipts_pending_dr"]
        )
        liabilities = money(
            s["contract_liabilities"] + s["co_ownership"] + s["receipts_pending"]
            + s["related_party_cr"] + s["borrowings_cr"] + s["tax_cr"]
            + s["bank_diff_cr"] + s["interaccount_cr"]
        )
        equity = money(s["share_capital"] + s["retained_earnings"])
        diff = money(assets - liabilities - equity)
        s["total_assets"] = assets
        s["total_liabilities"] = liabilities
        s["total_equity"] = equity
        s["balance_diff"] = diff
        checks.append({"year": y, "assets": assets, "liab_eq": money(liabilities + equity), "diff": diff})

    # IFRS identity on the sales-ledger population (excludes additional bank receipts,
    # which are added entirely to contract liabilities)
    ifrs_checks = []
    for y in years:
        col = collections[y]
        d_rec = money(rec_total[y] - (rec_total[y - 1] if y > 2023 else 0))
        d_cl_contracts = money(
            contracts[f"cl_{y}"].sum() - (contracts[f"cl_{y-1}"].sum() if y > 2023 else 0)
        )
        rev = money(contracts[f"rev_{y}"].sum())
        identity = money(col + d_rec - d_cl_contracts - rev)
        ifrs_checks.append({"year": y, "collections": col, "rev": rev, "d_rec": d_rec,
                             "d_cl": d_cl_contracts, "identity_diff": identity})

    sales_total = money(sales["amt"].sum())
    tb = {}
    for y in years:
        tb[y] = {}
        for acct in L.accounts():
            d, c = L.activity(acct, y, upto=True)
            net = money(d - c)
            if abs(net) >= 0.5:
                tb[y][acct] = net
    model = {
        "sfp": sfp,
        "pnl": pnl,
        "tax": tax,
        "ecl": ecl_by_year,
        "cash_proof": cash_proof,
        "stated_cash": stated,
        "opening_cash": OPENING_CASH,
        "opening_cash_total": OPENING_CASH_TOTAL,
        "share_capital": SHARE_CAPITAL,
        "collections": collections,
        "untraced_sales": untraced,
        "traced_sales_amt": money(traced_amt),
        "sales_total": sales_total,
        "addl_bank_receipts": addl_by_year,
        "cl_contracts": {y: money(contracts[f"cl_{y}"].sum()) for y in years},
        "cl_total": cl_total,
        "rec_total": rec_total,
        "rev_contracts": {y: money(contracts[f"rev_{y}"].sum()) for y in years},
        "ppe_add": ppe_add,
        "dep": dep_by_year,
        "cos_by_bucket": cos_by_year,
        "inv_close": inv_close,
        "inv_add": {y: dict(inv_add[y]) for y in years},
        "checks": checks,
        "ifrs_checks": ifrs_checks,
        "category_totals": cat,
        "contracts": contracts,
        "unallocated_in": un_in,
        "unallocated_out": un_out,
        "bank": bank,
        "sales": sales,
        "excl": excl,
        "exceptions_amt": money(exceptions["amt"].sum()) if len(exceptions) else 0,
        "exceptions_n": len(exceptions),
        "tb": tb,
        "n_customers": int(cust["Customer ID"].nunique()),
        "n_subs": len(subs),
        "n_sales": len(sales),
        "excl_total": money(excl["amt"].sum()),
    }
    return model


def print_summary(model):
    print("\n================ FINANCIAL SUMMARY ================")
    print(f"Sales ledger {model['sales_total']:,.2f}  traced {model['traced_sales_amt']:,.2f}")
    print("IFRS identity (should be ~0):")
    for row in model["ifrs_checks"]:
        print(" ", row)
    print("Balance check:")
    for row in model["checks"]:
        print(" ", row)
    presented = {
        "Cash", "Trade receivables", "ECL allowance", "Untraced customer receipts", "Untraced other receipts",
        "Inventory", "PPE cost", "Accumulated depreciation", "Payments pending allocation",
        "FX and other advances", "Inter-account transfers", "Related party current account",
        "Borrowings", "Bank statement difference", "Current tax", "Contract liabilities",
        "Contract liabilities - co-ownership", "Receipts pending allocation",
        "Share capital", "Retained earnings", "Dividends",
    }
    print("Unpresented TB balances:")
    for y, accts in model["tb"].items():
        for acct, net in accts.items():
            if acct not in presented and not acct.startswith("CAT::") and acct != "Opening equity clearing":
                print(f"  {y} {acct}: {net:,.2f}")
        cat_res = {a: n for a, n in accts.items() if a.startswith("CAT::") or a == "Opening equity clearing"}
        if cat_res:
            print(f"  {y} residuals {cat_res}")
    print("Cash proof:")
    for y, row in model["cash_proof"].items():
        print(f"  {y} ledger {row['ledger']:,.2f} stated {row['stated']:,.2f} diff {row['diff']:,.2f}")
    for y in (2023, 2024, 2025):
        p = model["pnl"][y]
        s = model["sfp"][y]
        print(f"\n--- {y} ---")
        print(f"  collections {model['collections'][y]:,.0f}  untraced {model['untraced_sales'][y]:,.0f}  addl bank {model['addl_bank_receipts'][y]:,.0f}")
        print(f"  revenue {p['revenue']:,.0f}  (land {p['rev_land']:,.0f} housing {p['rev_hous']:,.0f})")
        print(f"  other income {p['other_income']:,.0f}  COS {p['cos']:,.0f}  GP {p['gross_profit']:,.0f}")
        print(f"  opex {p['opex']:,.0f}  PBT {p['pbt']:,.0f}  tax {p['tax']:,.0f}  PAT {p['pat']:,.0f}")
        print(f"  cash {s['cash']:,.0f}  rec {s['receivables_net']:,.0f}  CL {s['contract_liabilities']:,.0f}")
        print(f"  inventory {s['inventory']:,.0f}  pending pay {s['payments_pending']:,.0f}  pending rec {s['receipts_pending']:,.0f}")
        print(f"  assets {s['total_assets']:,.0f}  L+E {s['total_liabilities']+s['total_equity']:,.0f}  diff {s['balance_diff']:,.2f}")
        print(f"  tax size {model['tax'][y]['size']} current {model['tax'][y]['current']:,.0f}")
    print("\nCategory coverage (naira):")
    cat = model["category_totals"]
    gross = model["bank"]["cr"].sum() + model["bank"]["dr"].sum()
    print(cat.groupby("category")[["dr", "cr"]].sum().sort_values("cr", ascending=False).to_string())
    un_in = model["bank"].loc[model["bank"]["category"] == "UNALLOCATED_IN", "cr"].sum()
    un_out = model["bank"].loc[model["bank"]["category"] == "UNALLOCATED_OUT", "dr"].sum()
    print(f"\nUnallocated in {un_in:,.0f} ({un_in/model['bank']['cr'].sum():.1%} of credits)")
    print(f"Unallocated out {un_out:,.0f} ({un_out/model['bank']['dr'].sum():.1%} of debits)")
    print("Top unallocated credits:")
    for _, r in model["unallocated_in"].head(12).iterrows():
        print(f"  {str(r['date'])[:10]} {r['cr']:,.0f} {str(r['narr'])[:110]}")
    print("Top unallocated debits:")
    for _, r in model["unallocated_out"].head(12).iterrows():
        print(f"  {str(r['date'])[:10]} {r['dr']:,.0f} {str(r['narr'])[:110]}")


def main():
    model = build()
    # drop heavy frames from the printed path but keep them in the pickle
    print_summary(model)
    with open("/tmp/baay_extract/model.pkl", "wb") as f:
        pickle.dump(model, f)
    print("\nSaved /tmp/baay_extract/model.pkl")


if __name__ == "__main__":
    main()
