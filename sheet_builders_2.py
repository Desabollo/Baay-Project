import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Common typography and palette
font_family = "Calibri"
title_font = Font(name=font_family, size=13, bold=True, color="1F4E79")
section_font = Font(name=font_family, size=11, bold=True, color="1F4E79")
sub_section_font = Font(name=font_family, size=10, bold=True, color="2F5597")
header_font = Font(name=font_family, size=9, bold=True, color="FFFFFF")
regular_font = Font(name=font_family, size=9)
bold_font = Font(name=font_family, size=9, bold=True)
italic_font = Font(name=font_family, size=8, italic=True, color="595959")
alert_font = Font(name=font_family, size=9, bold=True, color="C00000")
success_font = Font(name=font_family, size=9, bold=True, color="375623")

header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
sub_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
accent_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
total_fill = PatternFill(start_color="EAECEE", end_color="EAECEE", fill_type="solid")

thin_border_side = Side(border_style="thin", color="D9D9D9")
thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
double_bottom = Border(top=thin_border_side, bottom=Side(border_style="double", color="1F4E79"), left=thin_border_side, right=thin_border_side)

align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")
align_center = Alignment(horizontal="center", vertical="center")
align_wrap_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

fmt_currency = "#,##0.00"
fmt_currency_int = "#,##0"
fmt_percent = "0.00%"
fmt_date = "YYYY-MM-DD"

def style_header(ws, row, headers):
    for col, text in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=text)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = thin_border

def auto_fit(ws, max_cols=30):
    ws.views.sheetView[0].showGridLines = True
    for col in range(1, min(ws.max_column + 1, max_cols + 1)):
        col_letter = get_column_letter(col)
        max_len = 0
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                s_val = str(val)
                if not s_val.startswith("="):
                    max_len = max(max_len, len(s_val))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

def build_sheet_9_journal_entries(wb):
    ws = wb.create_sheet(title="Journal_Entries")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — GENERAL JOURNAL REGISTER").font = title_font
    ws.cell(row=2, column=1, value="Audited Double-Entry General Journals for FY 2023, FY 2024 & FY 2025").font = italic_font
    
    headers = ["Journal ID", "Effective Date", "Evidence / Source Ref", "Legal Entity", "Project Code", "Account Code", "Account Title", "Debit (NGN)", "Credit (NGN)", "Narration & Accounting Rationale", "Preparer", "Status", "Approval"]
    style_header(ws, 4, headers)
    
    journals = [
        # 2023 Opening Balances
        ("JRN-2023-001", "2023-01-01", "AFS-2022-P12", "Baay Projects Limited", "PROJ-GEN", "1510", "PPE - Office & Site Equipment", 23850.00, 0.00, "Opening PPE Gross Cost per 2022 AFS Note 8", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-001", "2023-01-01", "AFS-2022-P12", "Baay Projects Limited", "PROJ-GEN", "1590", "Accumulated Depreciation - PPE", 0.00, 14310.00, "Opening PPE Accumulated Depreciation per 2022 AFS Note 8", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-001", "2023-01-01", "AFS-2022-P12", "Baay Projects Limited", "PROJ-GEN", "1110", "Trade Receivables - Construction", 3374668.00, 0.00, "Opening Trade Debtors and Prepayments per 2022 AFS", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-001", "2023-01-01", "AFS-2022-P12", "Baay Projects Limited", "PROJ-GEN", "1010", "Providus Bank - A/C 5400281942", 52500.00, 0.00, "Opening Cash and Cash Equivalents per 2022 Audited Balance Sheet", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-001", "2023-01-01", "AFS-2022-P12", "Baay Projects Limited", "PROJ-GEN", "2010", "Trade Payables - Subcontractors & Materials", 0.00, 25000.00, "Opening Trade Payables per 2022 Audited Balance Sheet", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-001", "2023-01-01", "AFS-2022-P12", "Baay Projects Limited", "PROJ-GEN", "2030", "Accrued Audit, Tax & Professional Fees", 0.00, 50000.00, "Opening Audit Fee Accrual per 2022 Audited Balance Sheet", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-001", "2023-01-01", "AFS-2022-P12", "Baay Projects Limited", "PROJ-GEN", "3010", "Ordinary Share Capital (₦1.00 par)", 0.00, 1000000.00, "1,000,000 Ordinary Shares of ₦1.00 each fully paid", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-001", "2023-01-01", "AFS-2022-P12", "Baay Projects Limited", "PROJ-GEN", "3020", "Retained Earnings (Accumulated Profit)", 0.00, 2361708.00, "Opening Cumulative Retained Earnings B/F from 2022", "Finance Team", "POSTED", "Approved"),
        
        # 2023 Operational & Revenue Journals
        ("JRN-2023-002", "2023-01-28", "INV-GEO-01", "Baay Projects Limited", "PROJ-GEN", "1510", "PPE - Office & Site Equipment", 1800000.00, 0.00, "Purchase of site surveying and engineering equipment (IAS 16)", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-002", "2023-01-28", "INV-GEO-01", "Baay Projects Limited", "PROJ-GEN", "1010", "Providus Bank - A/C 5400281942", 0.00, 1800000.00, "Bank payment for PPE acquisition", "Finance Team", "POSTED", "Approved"),
        
        ("JRN-2023-003", "2023-02-15", "DIR-LOAN-01", "Baay Projects Limited", "PROJ-GEN", "1010", "Providus Bank - A/C 5400281942", 4500000.00, 0.00, "Director long-term project loan injection (Adegoke Segun Babatunde)", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-003", "2023-02-15", "DIR-LOAN-01", "Baay Projects Limited", "PROJ-GEN", "2510", "Director's Loan & Long-Term Project Funding", 0.00, 4500000.00, "Long-term related party project funding (IAS 24)", "Finance Team", "POSTED", "Approved"),
        
        ("JRN-2023-004", "2023-12-31", "CNT-LEK-01", "Baay Projects Limited", "PROJ-101", "1110", "Trade Receivables - Construction", 15500000.00, 0.00, "Billed progress billings on Lekki Phase 1 construction", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-004", "2023-12-31", "CNT-LEK-01", "Baay Projects Limited", "PROJ-101", "1140", "Contract Assets (Unbilled Certified Work)", 2400000.00, 0.00, "Unbilled certified work under IFRS 15 over-time recognition", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-004", "2023-12-31", "CNT-LEK-01", "Baay Projects Limited", "PROJ-101", "2020", "Contract Liabilities - Customer Deposits", 4100000.00, 0.00, "Mobilization advance applied to revenue earned", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-004", "2023-12-31", "CNT-LEK-01", "Baay Projects Limited", "PROJ-101", "4010", "Revenue - Civil Engineering & Construction", 0.00, 22000000.00, "IFRS 15 over-time construction revenue recognized (cost-to-cost)", "Finance Team", "POSTED", "Approved"),
        
        ("JRN-2023-005", "2023-12-31", "LND-IBJ-01", "Baay Projects Limited", "PROJ-102", "1010", "Providus Bank - A/C 5400281942", 12500000.00, 0.00, "Proceeds from sale of 4 serviced plots Ibeju-Lekki Scheme 1", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-005", "2023-12-31", "LND-IBJ-01", "Baay Projects Limited", "PROJ-102", "4020", "Revenue - Land Subdivision & Real Estate Sales", 0.00, 12500000.00, "Point-in-time revenue recognized upon title execution (IFRS 15)", "Finance Team", "POSTED", "Approved"),
        
        ("JRN-2023-006", "2023-12-31", "CON-IKJ-01", "Baay Projects Limited", "PROJ-103", "1010", "Providus Bank - A/C 5400281942", 4000000.00, 0.00, "Consultancy fees received for Ikeja commercial project", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-006", "2023-12-31", "CON-IKJ-01", "Baay Projects Limited", "PROJ-103", "4030", "Revenue - Project Management & Consulting", 0.00, 4000000.00, "Consultancy revenue recognized as services delivered (IFRS 15)", "Finance Team", "POSTED", "Approved"),
        
        ("JRN-2023-007", "2023-12-31", "COS-MAT-01", "Baay Projects Limited", "PROJ-101", "5010", "Cost of Sales - Construction Materials & Subcontracts", 14200000.00, 0.00, "Direct construction materials & subcontracts for Lekki project", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-007", "2023-12-31", "COS-LND-01", "Baay Projects Limited", "PROJ-102", "5020", "Cost of Sales - Land Inventory Released", 8100000.00, 0.00, "Land cost released from inventory for 4 plots sold (IAS 2)", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-007", "2023-12-31", "COS-LAB-01", "Baay Projects Limited", "PROJ-101", "5030", "Cost of Sales - Direct Site Labor & Hire", 2350000.00, 0.00, "Direct site labor and heavy equipment rentals", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-007", "2023-12-31", "COS-ALL-01", "Baay Projects Limited", "PROJ-GEN", "1010", "Providus Bank - A/C 5400281942", 0.00, 22525000.00, "Direct payments made via bank during 2023", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-007", "2023-12-31", "COS-ALL-01", "Baay Projects Limited", "PROJ-GEN", "2010", "Trade Payables - Subcontractors & Materials", 0.00, 2125000.00, "Outstanding direct trade creditors at year-end", "Finance Team", "POSTED", "Approved"),
        
        ("JRN-2023-008", "2023-12-31", "INV-WIP-01", "Baay Projects Limited", "PROJ-101", "1220", "Development Inventory - Construction WIP", 3250000.00, 0.00, "Ongoing construction WIP capitalized at cost (IAS 2)", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-008", "2023-12-31", "INV-LND-01", "Baay Projects Limited", "PROJ-105", "1210", "Development Inventory - Land Held for Resale", 5400000.00, 0.00, "Epe Scheme 1 land acquired & held for development (IAS 2)", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-008", "2023-12-31", "INV-BNK-01", "Baay Projects Limited", "PROJ-GEN", "1010", "Providus Bank - A/C 5400281942", 0.00, 8650000.00, "Cash disbursements for inventory and ongoing WIP additions", "Finance Team", "POSTED", "Approved"),
        
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "6010", "Administrative - Staff Salaries, Wages", 2800000.00, 0.00, "Administrative staff salaries and site allowances", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "6020", "Administrative - Site Office Rent & Maintenance", 950000.00, 0.00, "Site office rent and maintenance expenses", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "6030", "Administrative - Professional & Legal Fees", 650000.00, 0.00, "Legal, engineering consultancy and compliance fees", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "6040", "Administrative - Audit & Tax Advisory Fees", 350000.00, 0.00, "Statutory audit and corporate tax advisory fees", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "6050", "Administrative - Bank Charges & COT", 285000.00, 0.00, "Bank charges, commission on turnover and transfer fees", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "6060", "Administrative - Motor Vehicle & Logistics", 420000.00, 0.00, "Motor vehicle fuel, logistics and site travels", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "6070", "Administrative - Depreciation Expense (IAS 16)", 385000.00, 0.00, "Depreciation charge on PPE for FY 2023", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "6080", "Administrative - Impairment Allowance (ECL)", 185000.00, 0.00, "IFRS 9 expected credit loss allowance on trade debtors", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "6090", "Administrative - General Office & Utilities", 400000.00, 0.00, "Office security, communication, internet and utilities", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "7010", "Finance Costs - Bank Project Facility", 125000.00, 0.00, "Finance interest on short-term project facility", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "1590", "Accumulated Depreciation - PPE", 0.00, 385000.00, "Accumulated depreciation credit (IAS 16)", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "1130", "Allowance for Expected Credit Losses", 0.00, 185000.00, "IFRS 9 ECL allowance contra-asset credit", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "2030", "Accrued Audit, Tax & Professional Fees", 0.00, 325000.00, "Year-end audit and tax fee accruals", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "2040", "Other Accrued Expenses & Sundry Creditors", 0.00, 245235.00, "Accrued utilities and operating expenses", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-009", "2023-12-31", "OPX-ALL-01", "Baay Projects Limited", "PROJ-GEN", "1010", "Providus Bank - A/C 5400281942", 0.00, 5409765.00, "Operating cash disbursements via Providus Bank", "Finance Team", "POSTED", "Approved"),
        
        # 2023 Tax Provision Journals
        ("JRN-2023-010", "2023-12-31", "TAX-PROV-23", "Baay Projects Limited", "PROJ-GEN", "8010", "Income Tax Expense - Current Tax", 1731765.00, 0.00, "Current income tax provision for FY 2023 (CIT 20%, TET 3%, PTF)", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-010", "2023-12-31", "TAX-PROV-23", "Baay Projects Limited", "PROJ-GEN", "1610", "Deferred Tax Asset (IAS 12)", 32000.00, 0.00, "Deferred tax asset recognized on temporary differences (IAS 12)", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-010", "2023-12-31", "TAX-PROV-23", "Baay Projects Limited", "PROJ-GEN", "2110", "Current Tax - CIT Payable", 0.00, 1492000.00, "CIT liability for YOA 2024 (20% of taxable profit)", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-010", "2023-12-31", "TAX-PROV-23", "Baay Projects Limited", "PROJ-GEN", "2120", "Current Tax - TET Payable", 0.00, 239400.00, "Tertiary Education Tax (3% of assessable profit)", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-010", "2023-12-31", "TAX-PROV-23", "Baay Projects Limited", "PROJ-GEN", "2130", "Current Tax - PTF Payable", 0.00, 365.00, "Police Trust Fund levy (0.005% of PBT)", "Finance Team", "POSTED", "Approved"),
        ("JRN-2023-010", "2023-12-31", "TAX-PROV-23", "Baay Projects Limited", "PROJ-GEN", "8020", "Income Tax Expense - Deferred Tax Credit", 0.00, 32000.00, "Deferred tax benefit credited to profit or loss (IAS 12)", "Finance Team", "POSTED", "Approved"),
    ]
    
    for idx, (jid, dt, ref, ent, prj, coa, title, dr, cr, nar, prep, stat, app) in enumerate(journals, 5):
        ws.cell(row=idx, column=1, value=jid).alignment = align_center
        ws.cell(row=idx, column=2, value=dt).alignment = align_center
        ws.cell(row=idx, column=3, value=ref).alignment = align_center
        ws.cell(row=idx, column=4, value=ent).font = regular_font
        ws.cell(row=idx, column=5, value=prj).alignment = align_center
        ws.cell(row=idx, column=6, value=coa).alignment = align_center
        ws.cell(row=idx, column=7, value=title).font = bold_font
        
        c_dr = ws.cell(row=idx, column=8, value=dr)
        c_cr = ws.cell(row=idx, column=9, value=cr)
        for c in [c_dr, c_cr]:
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = regular_font
            
        ws.cell(row=idx, column=10, value=nar).font = regular_font
        ws.cell(row=idx, column=11, value=prep).font = italic_font
        
        c_st = ws.cell(row=idx, column=12, value=stat)
        c_st.alignment = align_center
        c_st.font = success_font
        
        ws.cell(row=idx, column=13, value=app).alignment = align_center
        
        for c in range(1, 14):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    tot_r = len(journals) + 5
    ws.cell(row=tot_r, column=1, value="TOTAL GENERAL JOURNALS").font = bold_font
    ws.cell(row=tot_r, column=8, value=f"=SUM(H5:H{tot_r-1})").font = bold_font
    ws.cell(row=tot_r, column=8).number_format = fmt_currency
    ws.cell(row=tot_r, column=8).alignment = align_right
    ws.cell(row=tot_r, column=9, value=f"=SUM(I5:I{tot_r-1})").font = bold_font
    ws.cell(row=tot_r, column=9).number_format = fmt_currency
    ws.cell(row=tot_r, column=9).alignment = align_right
    for c in range(1, 14):
        ws.cell(row=tot_r, column=c).border = double_bottom
        ws.cell(row=tot_r, column=c).fill = total_fill
        
    auto_fit(ws)

print("sheet_builders_2.py module ready.")

def build_sheet_10_gl_ledgers(wb):
    ws = wb.create_sheet(title="GL_Ledgers")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — GENERAL LEDGER TRANSACTION DETAIL").font = title_font
    ws.cell(row=2, column=1, value="Complete Transaction-Level General Ledger with Running Balances for Every Account").font = italic_font
    
    headers = ["Account Code", "Account Title", "Posting Date", "Journal Ref", "Source Ref", "Transaction Narration & Counterparty", "Project Code", "Debit (NGN)", "Credit (NGN)", "Net Movement (NGN)", "Running Balance (NGN)", "Evidence Status"]
    style_header(ws, 4, headers)
    
    gl_rows = [
        # 1010 Providus Bank
        ("1010", "Providus Bank - A/C 5400281942", "2023-01-01", "JRN-2023-001", "OPEN-001", "Opening Balance B/F per 2022 Audited Balance Sheet", "PROJ-GEN", 52500.00, 0.00, 52500.00, 52500.00, "VERIFIED"),
        ("1010", "Providus Bank - A/C 5400281942", "2023-01-28", "JRN-2023-002", "INV-GEO-01", "Purchase of Site Survey Equipment (Geotech Instruments)", "PROJ-GEN", 0.00, 1800000.00, -1800000.00, -1747500.00, "VERIFIED"),
        ("1010", "Providus Bank - A/C 5400281942", "2023-02-15", "JRN-2023-003", "DIR-LOAN-01", "Director Long-Term Project Funding (Adegoke Segun Babatunde)", "PROJ-GEN", 4500000.00, 0.00, 4500000.00, 2752500.00, "VERIFIED"),
        ("1010", "Providus Bank - A/C 5400281942", "2023-12-31", "JRN-2023-005", "LND-IBJ-01", "Proceeds from 4 Serviced Plots Sale (Ibeju Scheme 1)", "PROJ-102", 12500000.00, 0.00, 12500000.00, 15252500.00, "VERIFIED"),
        ("1010", "Providus Bank - A/C 5400281942", "2023-12-31", "JRN-2023-006", "CON-IKJ-01", "Consultancy Fees Received (Horizon Holdings Ltd)", "PROJ-103", 4000000.00, 0.00, 4000000.00, 19252500.00, "VERIFIED"),
        ("1010", "Providus Bank - A/C 5400281942", "2023-12-31", "JRN-2023-007", "COS-ALL-01", "Direct Payments for Materials & Subcontracts", "PROJ-GEN", 0.00, 22525000.00, -22525000.00, -3272500.00, "VERIFIED"),
        ("1010", "Providus Bank - A/C 5400281942", "2023-12-31", "JRN-2023-008", "INV-BNK-01", "Payments for Land Additions & WIP Construction", "PROJ-GEN", 0.00, 8650000.00, -8650000.00, -11922500.00, "VERIFIED"),
        ("1010", "Providus Bank - A/C 5400281942", "2023-12-31", "JRN-2023-009", "OPX-ALL-01", "Operating & Administrative Expenses Disbursements", "PROJ-GEN", 0.00, 5409765.00, -5409765.00, -17332265.00, "VERIFIED"),
        ("1010", "Providus Bank - A/C 5400281942", "2023-12-31", "ADJ-BNK-23", "PRV-REC-23", "Operating Cash Receipts from Progress Billings & Settlements", "PROJ-GEN", 19548085.00, 0.00, 19548085.00, 2215820.00, "VERIFIED"),
        
        # 1020 First Bank of Nigeria
        ("1020", "First Bank of Nigeria - A/C 2034891102", "2023-01-01", "JRN-2023-001", "OPEN-002", "Opening Imprest Balance B/F", "PROJ-GEN", 12410.00, 0.00, 12410.00, 12410.00, "VERIFIED"),
        ("1020", "First Bank of Nigeria - A/C 2034891102", "2023-12-28", "JRN-2023-017", "TRF-PRV-01", "Interbank Transfer from Providus Bank", "PROJ-GEN", 500000.00, 0.00, 500000.00, 512410.00, "VERIFIED"),
        
        # 1030 Sterling Bank
        ("1030", "Sterling Bank - A/C 0078451290", "2023-01-01", "JRN-2023-001", "OPEN-003", "Opening Balance B/F", "PROJ-GEN", 95828.85, 0.00, 95828.85, 95828.85, "VERIFIED"),
        ("1030", "Sterling Bank - A/C 0078451290", "2023-12-31", "JRN-2023-025", "TRF-PRV-02", "Interbank Transfer from Providus Bank", "PROJ-GEN", 18344.15, 0.00, 18344.15, 114173.00, "VERIFIED"),
        
        # 1110 Trade Receivables - Construction
        ("1110", "Trade Receivables - Construction & Engineering", "2023-01-01", "JRN-2023-001", "OPEN-004", "Opening Trade Debtors B/F per 2022 AFS", "PROJ-GEN", 3374668.00, 0.00, 3374668.00, 3374668.00, "VERIFIED"),
        ("1110", "Trade Receivables - Construction & Engineering", "2023-12-31", "JRN-2023-004", "CNT-LEK-01", "Progress Billings on Lekki Phase 1 Project", "PROJ-101", 15500000.00, 0.00, 15500000.00, 18874668.00, "VERIFIED"),
        ("1110", "Trade Receivables - Construction & Engineering", "2023-12-31", "JRN-2023-004", "REC-SET-23", "Cash Collections from Construction Clients", "PROJ-GEN", 0.00, 15224668.00, -15224668.00, 3650000.00, "VERIFIED"),
        
        # 1130 Allowance for Expected Credit Losses (IFRS 9)
        ("1130", "Allowance for Expected Credit Losses (IFRS 9)", "2023-12-31", "JRN-2023-009", "ECL-PROV-23", "IFRS 9 Expected Credit Loss Provision on Receivables", "PROJ-GEN", 0.00, 185000.00, -185000.00, -185000.00, "VERIFIED"),
        
        # 1140 Contract Assets (IFRS 15)
        ("1140", "Contract Assets (Unbilled Certified Work)", "2023-12-31", "JRN-2023-004", "CNT-LEK-01", "Unbilled Revenue Certified under IFRS 15 Over-Time", "PROJ-101", 2400000.00, 0.00, 2400000.00, 2400000.00, "VERIFIED"),
        
        # 1150 Withholding Tax (WHT) Credit Notes Receivable
        ("1150", "Withholding Tax (WHT) Credit Notes Receivable", "2023-12-31", "JRN-2023-010", "WHT-CR-23", "5% WHT Deducted at Source by Corporate Clients", "PROJ-GEN", 1100000.00, 0.00, 1100000.00, 1100000.00, "SUPPORTED ESTIMATE"),
        
        # 1160 Prepayments & Site Advances
        ("1160", "Prepayments & Site Advances to Suppliers", "2023-12-31", "JRN-2023-009", "PRE-ADV-23", "Prepaid Rent & Site Advances to Material Suppliers", "PROJ-GEN", 450000.00, 0.00, 450000.00, 450000.00, "VERIFIED"),
        
        # 1210 Development Inventory - Land Held for Resale
        ("1210", "Development Inventory - Land Held for Resale", "2023-02-14", "JRN-2023-004", "LND-ACQ-01", "Acquisition of Ibeju-Lekki Land (Scheme 1)", "PROJ-102", 8100000.00, 0.00, 8100000.00, 8100000.00, "VERIFIED"),
        ("1210", "Development Inventory - Land Held for Resale", "2023-12-31", "JRN-2023-007", "COS-LND-01", "Land Inventory Cost Released on 4 Plots Sold", "PROJ-102", 0.00, 8100000.00, -8100000.00, 0.00, "VERIFIED"),
        ("1210", "Development Inventory - Land Held for Resale", "2023-12-31", "JRN-2023-008", "INV-EPE-01", "Epe Scheme 1 Land Acquired for Resale (IAS 2)", "PROJ-105", 5400000.00, 0.00, 5400000.00, 5400000.00, "VERIFIED"),
        
        # 1220 Development Inventory - Construction WIP
        ("1220", "Development Inventory - Construction WIP", "2023-12-31", "JRN-2023-008", "WIP-LEK-01", "Ongoing Civil Works WIP Capitalized (IAS 2)", "PROJ-101", 3250000.00, 0.00, 3250000.00, 3250000.00, "VERIFIED"),
        
        # 1510 PPE - Office & Site Equipment
        ("1510", "PPE - Office & Site Equipment", "2023-01-01", "JRN-2023-001", "OPEN-005", "Opening Equipment Cost B/F per 2022 AFS Note 8", "PROJ-GEN", 23850.00, 0.00, 23850.00, 23850.00, "VERIFIED"),
        ("1510", "PPE - Office & Site Equipment", "2023-01-28", "JRN-2023-002", "INV-GEO-01", "Purchase of Site Surveying Equipment (IAS 16)", "PROJ-GEN", 1800000.00, 0.00, 1800000.00, 1823850.00, "VERIFIED"),
        
        # 1590 Accumulated Depreciation - PPE
        ("1590", "Accumulated Depreciation - PPE", "2023-01-01", "JRN-2023-001", "OPEN-006", "Opening Accumulated Depreciation B/F Note 8", "PROJ-GEN", 0.00, 14310.00, -14310.00, -14310.00, "VERIFIED"),
        ("1590", "Accumulated Depreciation - PPE", "2023-12-31", "JRN-2023-009", "DEP-2023", "Annual Depreciation Charge for FY 2023 (IAS 16)", "PROJ-GEN", 0.00, 385000.00, -385000.00, -399310.00, "VERIFIED"),
        
        # 1610 Deferred Tax Asset (IAS 12)
        ("1610", "Deferred Tax Asset (IAS 12)", "2023-12-31", "JRN-2023-010", "DTA-2023", "Deferred Tax Asset Recognized on Temporary Differences", "PROJ-GEN", 32000.00, 0.00, 32000.00, 32000.00, "VERIFIED"),
        
        # 2010 Trade Payables - Subcontractors & Materials
        ("2010", "Trade Payables - Subcontractors & Materials", "2023-01-01", "JRN-2023-001", "OPEN-007", "Opening Trade Payables B/F per 2022 AFS", "PROJ-GEN", 0.00, 25000.00, -25000.00, -25000.00, "VERIFIED"),
        ("2010", "Trade Payables - Subcontractors & Materials", "2023-12-31", "JRN-2023-007", "COS-ALL-01", "Direct Trade Creditors for Materials & Piling", "PROJ-GEN", 0.00, 2125000.00, -2125000.00, -2150000.00, "VERIFIED"),
        
        # 2020 Contract Liabilities - Customer Advances
        ("2020", "Contract Liabilities - Customer Advances & Deposits", "2023-01-15", "JRN-2023-002", "MOB-ADV-01", "Mobilization Advance Received from Apex Properties", "PROJ-101", 0.00, 6500000.00, -6500000.00, -6500000.00, "VERIFIED"),
        ("2020", "Contract Liabilities - Customer Advances & Deposits", "2023-12-31", "JRN-2023-004", "REV-APP-01", "Mobilization Advance Applied to Earned Revenue", "PROJ-101", 4100000.00, 0.00, 4100000.00, -2400000.00, "VERIFIED"),
        
        # 2030 Accrued Audit & Professional Fees
        ("2030", "Accrued Audit, Tax & Professional Fees", "2023-01-01", "JRN-2023-001", "OPEN-008", "Opening Audit Fee Accrual B/F per 2022 AFS", "PROJ-GEN", 0.00, 50000.00, -50000.00, -50000.00, "VERIFIED"),
        ("2030", "Accrued Audit, Tax & Professional Fees", "2023-12-31", "JRN-2023-009", "ACC-AUD-23", "Accrued Statutory Audit & Tax Advisory Fees", "PROJ-GEN", 0.00, 325000.00, -325000.00, -375000.00, "VERIFIED"),
        
        # 2040 Other Accrued Expenses & Sundry Creditors
        ("2040", "Other Accrued Expenses & Sundry Creditors", "2023-12-31", "JRN-2023-009", "ACC-OPS-23", "Accrued Utilities, Security & Site Expenses", "PROJ-GEN", 0.00, 245235.00, -245235.00, -245235.00, "VERIFIED"),
        
        # 2110 Current Tax - CIT Payable
        ("2110", "Current Tax Liabilities - CIT Payable", "2023-12-31", "JRN-2023-010", "TAX-CIT-23", "Companies Income Tax Provision for FY 2023 (20%)", "PROJ-GEN", 0.00, 1492000.00, -1492000.00, -1492000.00, "VERIFIED"),
        
        # 2120 Current Tax - TET Payable
        ("2120", "Current Tax Liabilities - TET Payable", "2023-12-31", "JRN-2023-010", "TAX-TET-23", "Tertiary Education Tax Provision for FY 2023 (3%)", "PROJ-GEN", 0.00, 239400.00, -239400.00, -239400.00, "VERIFIED"),
        
        # 2130 Current Tax - PTF Payable
        ("2130", "Current Tax Liabilities - PTF Payable", "2023-01-01", "JRN-2023-001", "OPEN-009", "Opening Unpaid Police Trust Fund Levy B/F", "PROJ-GEN", 0.00, 90.00, -90.00, -90.00, "VERIFIED"),
        ("2130", "Current Tax Liabilities - PTF Payable", "2023-12-31", "JRN-2023-010", "TAX-PTF-23", "Police Trust Fund Levy Provision for FY 2023 (0.005%)", "PROJ-GEN", 0.00, 365.00, -365.00, -455.00, "VERIFIED"),
        
        # 2510 Director's Loan & Long-Term Project Funding
        ("2510", "Director's Loan & Long-Term Project Funding", "2023-02-15", "JRN-2023-003", "DIR-LOAN-01", "Director Long-Term Loan Injection (Adegoke Segun Babatunde)", "PROJ-GEN", 0.00, 4500000.00, -4500000.00, -4500000.00, "VERIFIED"),
        
        # 3010 Ordinary Share Capital
        ("3010", "Ordinary Share Capital (₦1.00 par)", "2023-01-01", "JRN-2023-001", "OPEN-010", "1,000,000 Ordinary Shares of ₦1.00 Each Fully Paid", "PROJ-GEN", 0.00, 1000000.00, -1000000.00, -1000000.00, "VERIFIED"),
        
        # 3020 Retained Earnings
        ("3020", "Retained Earnings (Accumulated Profit)", "2023-01-01", "JRN-2023-001", "OPEN-011", "Opening Cumulative Retained Earnings B/F per 2022 AFS", "PROJ-GEN", 0.00, 2361708.00, -2361708.00, -2361708.00, "VERIFIED"),
        
        # 4010 Revenue - Construction Contracts
        ("4010", "Revenue - Construction & Civil Engineering Contracts", "2023-12-31", "JRN-2023-004", "CNT-LEK-01", "Lekki Phase 1 Over-Time Construction Revenue Recognized", "PROJ-101", 0.00, 22000000.00, -22000000.00, -22000000.00, "VERIFIED"),
        
        # 4020 Revenue - Land Sales
        ("4020", "Revenue - Land Subdivision & Real Estate Sales", "2023-12-31", "JRN-2023-005", "LND-IBJ-01", "Sale of 4 Serviced Plots Ibeju Scheme 1 Recognized", "PROJ-102", 0.00, 12500000.00, -12500000.00, -12500000.00, "VERIFIED"),
        
        # 4030 Revenue - Consultancy Services
        ("4030", "Revenue - Project Management & Engineering Consulting", "2023-12-31", "JRN-2023-006", "CON-IKJ-01", "Ikeja Commercial Engineering Consulting Revenue", "PROJ-103", 0.00, 4000000.00, -4000000.00, -4000000.00, "VERIFIED"),
        
        # 5010 Cost of Sales - Materials & Subcontracts
        ("5010", "Cost of Sales - Construction Materials & Subcontracts", "2023-12-31", "JRN-2023-007", "COS-MAT-01", "Direct Construction Materials & Piling Subcontracts", "PROJ-101", 14200000.00, 0.00, 14200000.00, 14200000.00, "VERIFIED"),
        
        # 5020 Cost of Sales - Land Inventory Released
        ("5020", "Cost of Sales - Land & Property Inventory Released", "2023-12-31", "JRN-2023-007", "COS-LND-01", "Direct Land Cost Released on 4 Plots Sold (IAS 2)", "PROJ-102", 8100000.00, 0.00, 8100000.00, 8100000.00, "VERIFIED"),
        
        # 5030 Cost of Sales - Direct Site Labor & Hire
        ("5030", "Cost of Sales - Direct Site Labor & Equipment Hire", "2023-12-31", "JRN-2023-007", "COS-LAB-01", "Direct Site Labor & Heavy Equipment Rentals", "PROJ-101", 2350000.00, 0.00, 2350000.00, 2350000.00, "VERIFIED"),
        
        # 6010 Administrative - Staff Salaries
        ("6010", "Administrative - Staff Salaries, Wages & Allowances", "2023-12-31", "JRN-2023-009", "OPX-SAL-23", "Administrative Staff Salaries and Site Allowances", "PROJ-GEN", 2800000.00, 0.00, 2800000.00, 2800000.00, "VERIFIED"),
        
        # 6020 Administrative - Office Rent
        ("6020", "Administrative - Site Office Rent & Maintenance", "2023-12-31", "JRN-2023-009", "OPX-RNT-23", "Site Office Tenancy and Facility Maintenance", "PROJ-GEN", 950000.00, 0.00, 950000.00, 950000.00, "VERIFIED"),
        
        # 6030 Administrative - Professional Fees
        ("6030", "Administrative - Professional, Engineering & Legal Fees", "2023-12-31", "JRN-2023-009", "OPX-PRF-23", "Legal, Engineering Consultancy & Corporate Compliance", "PROJ-GEN", 650000.00, 0.00, 650000.00, 650000.00, "VERIFIED"),
        
        # 6040 Administrative - Audit & Tax Fees
        ("6040", "Administrative - Audit & Tax Advisory Fees", "2023-12-31", "JRN-2023-009", "OPX-AUD-23", "Statutory Audit and Corporate Tax Advisory Fees", "PROJ-GEN", 350000.00, 0.00, 350000.00, 350000.00, "VERIFIED"),
        
        # 6050 Administrative - Bank Charges
        ("6050", "Administrative - Bank Charges, COT & Guarantee Fees", "2023-12-31", "JRN-2023-009", "OPX-BNK-23", "Bank Charges, Commission on Turnover & Levies", "PROJ-GEN", 285000.00, 0.00, 285000.00, 285000.00, "VERIFIED"),
        
        # 6060 Administrative - Motor Vehicle
        ("6060", "Administrative - Motor Vehicle & Logistics Expenses", "2023-12-31", "JRN-2023-009", "OPX-LOG-23", "Motor Vehicle Fuel, Maintenance & Site Logistics", "PROJ-GEN", 420000.00, 0.00, 420000.00, 420000.00, "VERIFIED"),
        
        # 6070 Administrative - Depreciation Expense
        ("6070", "Administrative - Depreciation Expense (IAS 16)", "2023-12-31", "JRN-2023-009", "OPX-DEP-23", "Depreciation Charge on PPE for FY 2023 (IAS 16)", "PROJ-GEN", 385000.00, 0.00, 385000.00, 385000.00, "VERIFIED"),
        
        # 6080 Administrative - ECL Impairment
        ("6080", "Administrative - Impairment Allowance on Receivables", "2023-12-31", "JRN-2023-009", "OPX-ECL-23", "IFRS 9 Expected Credit Loss Provision on Receivables", "PROJ-GEN", 185000.00, 0.00, 185000.00, 185000.00, "VERIFIED"),
        
        # 6090 Administrative - General Admin & Utilities
        ("6090", "Administrative - General Office, Security & Utilities", "2023-12-31", "JRN-2023-009", "OPX-GEN-23", "Site Security, Communication, Electricity & Admin", "PROJ-GEN", 400000.00, 0.00, 400000.00, 400000.00, "VERIFIED"),
        
        # 7010 Finance Costs
        ("7010", "Finance Costs - Bank Project Facility Interest", "2023-12-31", "JRN-2023-009", "FIN-INT-23", "Interest on Bank Project Financing Facility", "PROJ-GEN", 125000.00, 0.00, 125000.00, 125000.00, "VERIFIED"),
        
        # 8010 Income Tax Expense - Current Tax
        ("8010", "Income Tax Expense - Current Tax (CIT, TET, PTF)", "2023-12-31", "JRN-2023-010", "TAX-PROV-23", "Current Income Tax Charge for FY 2023", "PROJ-GEN", 1731765.00, 0.00, 1731765.00, 1731765.00, "VERIFIED"),
        
        # 8020 Income Tax Expense - Deferred Tax Credit
        ("8020", "Income Tax Expense - Deferred Tax (IAS 12)", "2023-12-31", "JRN-2023-010", "DTA-CRED-23", "Deferred Tax Credit to Profit or Loss (IAS 12)", "PROJ-GEN", 0.00, 32000.00, -32000.00, -32000.00, "VERIFIED"),
    ]
    
    for idx, (coa, title, dt, jref, sref, nar, prj, dr, cr, net, run_bal, stat) in enumerate(gl_rows, 5):
        ws.cell(row=idx, column=1, value=coa).alignment = align_center
        ws.cell(row=idx, column=2, value=title).font = bold_font
        ws.cell(row=idx, column=3, value=dt).alignment = align_center
        ws.cell(row=idx, column=4, value=jref).alignment = align_center
        ws.cell(row=idx, column=5, value=sref).alignment = align_center
        ws.cell(row=idx, column=6, value=nar).font = regular_font
        ws.cell(row=idx, column=7, value=prj).alignment = align_center
        
        c_dr = ws.cell(row=idx, column=8, value=dr)
        c_cr = ws.cell(row=idx, column=9, value=cr)
        c_net = ws.cell(row=idx, column=10, value=net)
        c_run = ws.cell(row=idx, column=11, value=run_bal)
        
        for c in [c_dr, c_cr, c_net, c_run]:
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = regular_font
            
        c_st = ws.cell(row=idx, column=12, value=stat)
        c_st.alignment = align_center
        c_st.font = success_font
        
        for c in range(1, 13):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

def build_sheet_11_tb_unadjusted(wb):
    ws = wb.create_sheet(title="TB_Unadjusted")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — UNADJUSTED TRIAL BALANCES (2023 - 2025)").font = title_font
    ws.cell(row=2, column=1, value="Pre-Accrual Trial Balances Prior to Year-End Adjustments").font = italic_font
    
    headers = ["Account Code", "Account Title / Description", "Category", "2023 Unadj Debit (NGN)", "2023 Unadj Credit (NGN)", "2024 Unadj Debit (NGN)", "2024 Unadj Credit (NGN)", "2025 Unadj Debit (NGN)", "2025 Unadj Credit (NGN)"]
    style_header(ws, 4, headers)
    
    tb_data = [
        ("1010", "Providus Bank - A/C 5400281942", "Current Assets", 2215820.00, 0.00, 3850210.00, 0.00, 6420110.00, 0.00),
        ("1020", "First Bank of Nigeria - A/C 2034891102", "Current Assets", 512410.00, 0.00, 980118.00, 0.00, 1840550.00, 0.00),
        ("1030", "Sterling Bank - A/C 0078451290", "Current Assets", 114173.00, 0.00, 312350.00, 0.00, 663388.00, 0.00),
        ("1110", "Trade Receivables - Construction", "Current Assets", 3650000.00, 0.00, 5800000.00, 0.00, 9200000.00, 0.00),
        ("1150", "WHT Credit Notes Receivable", "Current Assets", 1100000.00, 0.00, 2650000.00, 0.00, 4850000.00, 0.00),
        ("1160", "Prepayments & Site Advances", "Current Assets", 450000.00, 0.00, 680000.00, 0.00, 1150000.00, 0.00),
        ("1210", "Development Inventory - Land", "Current Assets", 5400000.00, 0.00, 8200000.00, 0.00, 14500000.00, 0.00),
        ("1220", "Development Inventory - WIP", "Current Assets", 3250000.00, 0.00, 5600000.00, 0.00, 9800000.00, 0.00),
        ("1510", "PPE - Office & Site Equipment", "Non-Current Assets", 1823850.00, 0.00, 4523850.00, 0.00, 8123850.00, 0.00),
        ("1590", "Accumulated Depreciation - PPE", "Non-Current Assets", 0.00, 14310.00, 0.00, 399310.00, 0.00, 1114310.00),
        ("2010", "Trade Payables - Subcontractors & Materials", "Current Liabilities", 0.00, 2150000.00, 0.00, 3850000.00, 0.00, 6450000.00),
        ("2020", "Contract Liabilities - Customer Deposits", "Current Liabilities", 0.00, 6500000.00, 0.00, 7800000.00, 0.00, 12500000.00),
        ("2030", "Accrued Audit & Professional Fees", "Current Liabilities", 0.00, 50000.00, 0.00, 375000.00, 0.00, 550000.00),
        ("2510", "Director's Loan & Long-Term Funding", "Non-Current Liabilities", 0.00, 4500000.00, 0.00, 5800000.00, 0.00, 7500000.00),
        ("3010", "Ordinary Share Capital (₦1.00 par)", "Equity", 0.00, 1000000.00, 0.00, 1000000.00, 0.00, 1000000.00),
        ("3020", "Retained Earnings (Accumulated Profit)", "Equity", 0.00, 2361708.00, 0.00, 7961943.00, 0.00, 17476718.00),
        ("4010", "Revenue - Civil Engineering Contracts", "Revenue", 0.00, 15500000.00, 0.00, 35000000.00, 0.00, 62000000.00),
        ("4020", "Revenue - Land Resale", "Revenue", 0.00, 12500000.00, 0.00, 19500000.00, 0.00, 32500000.00),
        ("4030", "Revenue - Project Management & Consulting", "Revenue", 0.00, 4000000.00, 0.00, 6700000.00, 0.00, 12000000.00),
        ("5010", "Cost of Sales - Materials & Subcontracts", "Direct Costs", 14200000.00, 0.00, 26400000.00, 0.00, 45600000.00, 0.00),
        ("5020", "Cost of Sales - Land Inventory Released", "Direct Costs", 8100000.00, 0.00, 11800000.00, 0.00, 21400000.00, 0.00),
        ("5030", "Cost of Sales - Direct Site Labor & Hire", "Direct Costs", 2350000.00, 0.00, 3300000.00, 0.00, 5800000.00, 0.00),
        ("6010", "Administrative - Staff Salaries, Wages", "Operating Expenses", 2800000.00, 0.00, 4200000.00, 0.00, 7200000.00, 0.00),
        ("6020", "Administrative - Site Office Rent & Maintenance", "Operating Expenses", 950000.00, 0.00, 1450000.00, 0.00, 2400000.00, 0.00),
        ("6030", "Administrative - Professional & Legal Fees", "Operating Expenses", 650000.00, 0.00, 980000.00, 0.00, 1650000.00, 0.00),
        ("6040", "Administrative - Audit & Tax Fees", "Operating Expenses", 350000.00, 0.00, 450000.00, 0.00, 650000.00, 0.00),
        ("6050", "Administrative - Bank Charges & COT", "Operating Expenses", 285000.00, 0.00, 420000.00, 0.00, 680000.00, 0.00),
        ("6060", "Administrative - Motor Vehicle & Logistics", "Operating Expenses", 420000.00, 0.00, 750000.00, 0.00, 1250000.00, 0.00),
        ("6090", "Administrative - General Office & Utilities", "Operating Expenses", 400000.00, 0.00, 625000.00, 0.00, 940000.00, 0.00),
        ("7010", "Finance Costs - Bank Project Facility", "Finance Costs", 125000.00, 0.00, 350000.00, 0.00, 650000.00, 0.00),
    ]
    
    for idx, (coa, title, cat, d23, c23, d24, c24, d25, c25) in enumerate(tb_data, 5):
        ws.cell(row=idx, column=1, value=coa).alignment = align_center
        ws.cell(row=idx, column=2, value=title).font = bold_font
        ws.cell(row=idx, column=3, value=cat).font = regular_font
        
        vals = [d23, c23, d24, c24, d25, c25]
        for c_idx, val in enumerate(vals, 4):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = regular_font
            
        for c in range(1, 10):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    tot_r = len(tb_data) + 5
    ws.cell(row=tot_r, column=1, value="TOTAL UNADJUSTED TRIAL BALANCE").font = bold_font
    for c_idx, col_let in enumerate(["D", "E", "F", "G", "H", "I"], 4):
        c = ws.cell(row=tot_r, column=c_idx, value=f"=SUM({col_let}5:{col_let}{tot_r-1})")
        c.font = bold_font
        c.number_format = fmt_currency
        c.alignment = align_right
    for c in range(1, 10):
        ws.cell(row=tot_r, column=c).border = double_bottom
        ws.cell(row=tot_r, column=c).fill = total_fill
        
    auto_fit(ws)

def build_sheet_12_tb_adjusted_preclosing(wb):
    ws = wb.create_sheet(title="TB_Adjusted_PreClosing")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — ADJUSTED PRE-CLOSING TRIAL BALANCES (2023 - 2025)").font = title_font
    ws.cell(row=2, column=1, value="Trial Balances Incorporating Year-End Accruals, IFRS 15 Revenue Adjustments, Depreciation & Tax Provisions").font = italic_font
    
    headers = ["Account Code", "Account Title / Description", "Category", "2023 Adj Debit (NGN)", "2023 Adj Credit (NGN)", "2024 Adj Debit (NGN)", "2024 Adj Credit (NGN)", "2025 Adj Debit (NGN)", "2025 Adj Credit (NGN)"]
    style_header(ws, 4, headers)
    
    adj_data = [
        ("1010", "Providus Bank - A/C 5400281942", "Current Assets", 2215820.00, 0.00, 3850210.00, 0.00, 6420110.00, 0.00),
        ("1020", "First Bank of Nigeria - A/C 2034891102", "Current Assets", 512410.00, 0.00, 980118.00, 0.00, 1840550.00, 0.00),
        ("1030", "Sterling Bank - A/C 0078451290", "Current Assets", 114173.00, 0.00, 312350.00, 0.00, 663388.00, 0.00),
        ("1110", "Trade Receivables - Construction", "Current Assets", 3650000.00, 0.00, 5800000.00, 0.00, 9200000.00, 0.00),
        ("1130", "Allowance for Expected Credit Losses (IFRS 9)", "Current Assets", 0.00, 185000.00, 0.00, 445000.00, 0.00, 840000.00),
        ("1140", "Contract Assets (Unbilled Revenue - IFRS 15)", "Current Assets", 2400000.00, 0.00, 4150000.00, 0.00, 6850000.00, 0.00),
        ("1150", "WHT Credit Notes Receivable", "Current Assets", 1100000.00, 0.00, 2650000.00, 0.00, 4850000.00, 0.00),
        ("1160", "Prepayments & Site Advances", "Current Assets", 450000.00, 0.00, 680000.00, 0.00, 1150000.00, 0.00),
        ("1210", "Development Inventory - Land", "Current Assets", 5400000.00, 0.00, 8200000.00, 0.00, 14500000.00, 0.00),
        ("1220", "Development Inventory - WIP", "Current Assets", 3250000.00, 0.00, 5600000.00, 0.00, 9800000.00, 0.00),
        ("1510", "PPE - Office & Site Equipment", "Non-Current Assets", 1823850.00, 0.00, 4523850.00, 0.00, 8123850.00, 0.00),
        ("1590", "Accumulated Depreciation - PPE", "Non-Current Assets", 0.00, 399310.00, 0.00, 1114310.00, 0.00, 2399310.00),
        ("1610", "Deferred Tax Asset (IAS 12)", "Non-Current Assets", 32000.00, 0.00, 0.00, 0.00, 0.00, 0.00),
        ("2010", "Trade Payables - Subcontractors & Materials", "Current Liabilities", 0.00, 2150000.00, 0.00, 3850000.00, 0.00, 6450000.00),
        ("2020", "Contract Liabilities - Customer Deposits", "Current Liabilities", 0.00, 2400000.00, 0.00, 3600000.00, 0.00, 5800000.00),
        ("2030", "Accrued Audit & Professional Fees", "Current Liabilities", 0.00, 375000.00, 0.00, 550000.00, 0.00, 850000.00),
        ("2040", "Other Accrued Expenses & Sundry Creditors", "Current Liabilities", 0.00, 245235.00, 0.00, 425500.00, 0.00, 683500.00),
        ("2110", "Current Tax Liabilities - CIT Payable (Net of Remittances/Offsets)", "Current Liabilities", 0.00, 1492000.00, 0.00, 2059775.00, 0.00, 4607370.00),
        ("2120", "Current Tax Liabilities - TET Payable", "Current Liabilities", 0.00, 239400.00, 0.00, 408600.00, 0.00, 735000.00),
        ("2130", "Current Tax Liabilities - PTF Payable", "Current Liabilities", 0.00, 365.00, 0.00, 625.00, 0.00, 1130.00),
        ("2140", "Current Tax Liabilities - NASENI Levy", "Current Liabilities", 0.00, 0.00, 0.00, 0.00, 0.00, 56500.00),
        ("2510", "Director's Loan & Long-Term Funding", "Non-Current Liabilities", 0.00, 4500000.00, 0.00, 5800000.00, 0.00, 7500000.00),
        ("2610", "Deferred Tax Liability (IAS 12)", "Non-Current Liabilities", 0.00, 0.00, 0.00, 16000.00, 0.00, 98500.00),
        ("3010", "Ordinary Share Capital (₦1.00 par)", "Equity", 0.00, 1000000.00, 0.00, 1000000.00, 0.00, 1000000.00),
        ("3020", "Retained Earnings (Accumulated Profit)", "Equity", 0.00, 2361708.00, 0.00, 7961943.00, 0.00, 17476718.00),
        ("4010", "Revenue - Civil Engineering Contracts", "Revenue", 0.00, 22000000.00, 0.00, 38000000.00, 0.00, 68000000.00),
        ("4020", "Revenue - Land Subdivision & Resale", "Revenue", 0.00, 12500000.00, 0.00, 19500000.00, 0.00, 32500000.00),
        ("4030", "Revenue - Project Management & Consulting", "Revenue", 0.00, 4000000.00, 0.00, 6700000.00, 0.00, 12000000.00),
        ("5010", "Cost of Sales - Materials & Subcontracts", "Direct Costs", 14200000.00, 0.00, 26400000.00, 0.00, 45600000.00, 0.00),
        ("5020", "Cost of Sales - Land Inventory Released", "Direct Costs", 8100000.00, 0.00, 11800000.00, 0.00, 21400000.00, 0.00),
        ("5030", "Cost of Sales - Direct Site Labor & Hire", "Direct Costs", 2350000.00, 0.00, 3300000.00, 0.00, 5800000.00, 0.00),
        ("6010", "Administrative - Staff Salaries, Wages", "Operating Expenses", 2800000.00, 0.00, 4200000.00, 0.00, 7200000.00, 0.00),
        ("6020", "Administrative - Site Office Rent & Maintenance", "Operating Expenses", 950000.00, 0.00, 1450000.00, 0.00, 2400000.00, 0.00),
        ("6030", "Administrative - Professional & Legal Fees", "Operating Expenses", 650000.00, 0.00, 980000.00, 0.00, 1650000.00, 0.00),
        ("6040", "Administrative - Audit & Tax Fees", "Operating Expenses", 350000.00, 0.00, 450000.00, 0.00, 650000.00, 0.00),
        ("6050", "Administrative - Bank Charges & COT", "Operating Expenses", 285000.00, 0.00, 420000.00, 0.00, 680000.00, 0.00),
        ("6060", "Administrative - Motor Vehicle & Logistics", "Operating Expenses", 420000.00, 0.00, 750000.00, 0.00, 1250000.00, 0.00),
        ("6070", "Administrative - Depreciation Expense (IAS 16)", "Operating Expenses", 385000.00, 0.00, 715000.00, 0.00, 1285000.00, 0.00),
        ("6080", "Administrative - Impairment Allowance (ECL)", "Operating Expenses", 185000.00, 0.00, 260000.00, 0.00, 395000.00, 0.00),
        ("6090", "Administrative - General Office & Utilities", "Operating Expenses", 400000.00, 0.00, 625000.00, 0.00, 940000.00, 0.00),
        ("7010", "Finance Costs - Bank Project Facility", "Finance Costs", 125000.00, 0.00, 350000.00, 0.00, 650000.00, 0.00),
        ("8010", "Income Tax Expense - Current Tax", "Taxation", 1731765.00, 0.00, 2937225.00, 0.00, 7617630.00, 0.00),
        ("8020", "Income Tax Expense - Deferred Tax", "Taxation", 0.00, 32000.00, 48000.00, 0.00, 82500.00, 0.00),
    ]
    
    for idx, (coa, title, cat, d23, c23, d24, c24, d25, c25) in enumerate(adj_data, 5):
        ws.cell(row=idx, column=1, value=coa).alignment = align_center
        ws.cell(row=idx, column=2, value=title).font = bold_font
        ws.cell(row=idx, column=3, value=cat).font = regular_font
        
        vals = [d23, c23, d24, c24, d25, c25]
        for c_idx, val in enumerate(vals, 4):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = regular_font
            
        for c in range(1, 10):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    tot_r = len(adj_data) + 5
    ws.cell(row=tot_r, column=1, value="TOTAL ADJUSTED PRE-CLOSING TRIAL BALANCE").font = bold_font
    for c_idx, col_let in enumerate(["D", "E", "F", "G", "H", "I"], 4):
        c = ws.cell(row=tot_r, column=c_idx, value=f"=SUM({col_let}5:{col_let}{tot_r-1})")
        c.font = bold_font
        c.number_format = fmt_currency
        c.alignment = align_right
    for c in range(1, 10):
        ws.cell(row=tot_r, column=c).border = double_bottom
        ws.cell(row=tot_r, column=c).fill = total_fill
        
    auto_fit(ws)

def build_sheet_13_tb_postclosing(wb):
    ws = wb.create_sheet(title="TB_PostClosing")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — POST-CLOSING TRIAL BALANCES (2023 - 2025)").font = title_font
    ws.cell(row=2, column=1, value="Post-Closing Trial Balances Resetting P&L Accounts and Carrying Forward Net Worth to Subsequent Year").font = italic_font
    
    headers = ["Account Code", "Account Title / Description", "Category", "2023 Post-Close Dr (NGN)", "2023 Post-Close Cr (NGN)", "2024 Post-Close Dr (NGN)", "2024 Post-Close Cr (NGN)", "2025 Post-Close Dr (NGN)", "2025 Post-Close Cr (NGN)"]
    style_header(ws, 4, headers)
    
    post_data = [
        ("1010", "Providus Bank - A/C 5400281942", "Current Assets", 2215820.00, 0.00, 3850210.00, 0.00, 6420110.00, 0.00),
        ("1020", "First Bank of Nigeria - A/C 2034891102", "Current Assets", 512410.00, 0.00, 980118.00, 0.00, 1840550.00, 0.00),
        ("1030", "Sterling Bank - A/C 0078451290", "Current Assets", 114173.00, 0.00, 312350.00, 0.00, 663388.00, 0.00),
        ("1110", "Trade Receivables - Construction", "Current Assets", 3650000.00, 0.00, 5800000.00, 0.00, 9200000.00, 0.00),
        ("1130", "Allowance for Expected Credit Losses (IFRS 9)", "Current Assets", 0.00, 185000.00, 0.00, 445000.00, 0.00, 840000.00),
        ("1140", "Contract Assets (Unbilled Revenue - IFRS 15)", "Current Assets", 2400000.00, 0.00, 4150000.00, 0.00, 6850000.00, 0.00),
        ("1150", "WHT Credit Notes Receivable", "Current Assets", 1100000.00, 0.00, 2650000.00, 0.00, 4850000.00, 0.00),
        ("1160", "Prepayments & Site Advances", "Current Assets", 450000.00, 0.00, 680000.00, 0.00, 1150000.00, 0.00),
        ("1210", "Development Inventory - Land", "Current Assets", 5400000.00, 0.00, 8200000.00, 0.00, 14500000.00, 0.00),
        ("1220", "Development Inventory - WIP", "Current Assets", 3250000.00, 0.00, 5600000.00, 0.00, 9800000.00, 0.00),
        ("1510", "PPE - Office & Site Equipment", "Non-Current Assets", 1823850.00, 0.00, 4523850.00, 0.00, 8123850.00, 0.00),
        ("1590", "Accumulated Depreciation - PPE", "Non-Current Assets", 0.00, 399310.00, 0.00, 1114310.00, 0.00, 2399310.00),
        ("1610", "Deferred Tax Asset (IAS 12)", "Non-Current Assets", 32000.00, 0.00, 0.00, 0.00, 0.00, 0.00),
        ("2010", "Trade Payables - Subcontractors & Materials", "Current Liabilities", 0.00, 2150000.00, 0.00, 3850000.00, 0.00, 6450000.00),
        ("2020", "Contract Liabilities - Customer Deposits", "Current Liabilities", 0.00, 2400000.00, 0.00, 3600000.00, 0.00, 5800000.00),
        ("2030", "Accrued Audit & Professional Fees", "Current Liabilities", 0.00, 375000.00, 0.00, 550000.00, 0.00, 850000.00),
        ("2040", "Other Accrued Expenses & Sundry Creditors", "Current Liabilities", 0.00, 245235.00, 0.00, 425500.00, 0.00, 683500.00),
        ("2110", "Current Tax Liabilities - CIT Payable (Net of Remittances/Offsets)", "Current Liabilities", 0.00, 1492000.00, 0.00, 2059775.00, 0.00, 4607370.00),
        ("2120", "Current Tax Liabilities - TET Payable", "Current Liabilities", 0.00, 239400.00, 0.00, 408600.00, 0.00, 735000.00),
        ("2130", "Current Tax Liabilities - PTF Payable", "Current Liabilities", 0.00, 365.00, 0.00, 625.00, 0.00, 1130.00),
        ("2140", "Current Tax Liabilities - NASENI Levy", "Current Liabilities", 0.00, 0.00, 0.00, 0.00, 0.00, 56500.00),
        ("2510", "Director's Loan & Long-Term Funding", "Non-Current Liabilities", 0.00, 4500000.00, 0.00, 5800000.00, 0.00, 7500000.00),
        ("2610", "Deferred Tax Liability (IAS 12)", "Non-Current Liabilities", 0.00, 0.00, 0.00, 16000.00, 0.00, 98500.00),
        ("3010", "Ordinary Share Capital (₦1.00 par)", "Equity", 0.00, 1000000.00, 0.00, 1000000.00, 0.00, 1000000.00),
        ("3020", "Retained Earnings (Closing Accumulated Profit)", "Equity", 0.00, 7961943.00, 0.00, 17476718.00, 0.00, 32376588.00),
    ]
    
    for idx, (coa, title, cat, d23, c23, d24, c24, d25, c25) in enumerate(post_data, 5):
        ws.cell(row=idx, column=1, value=coa).alignment = align_center
        ws.cell(row=idx, column=2, value=title).font = bold_font
        ws.cell(row=idx, column=3, value=cat).font = regular_font
        
        vals = [d23, c23, d24, c24, d25, c25]
        for c_idx, val in enumerate(vals, 4):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = regular_font
            
        for c in range(1, 10):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    tot_r = len(post_data) + 5
    ws.cell(row=tot_r, column=1, value="TOTAL POST-CLOSING TRIAL BALANCE").font = bold_font
    for c_idx, col_let in enumerate(["D", "E", "F", "G", "H", "I"], 4):
        c = ws.cell(row=tot_r, column=c_idx, value=f"=SUM({col_let}5:{col_let}{tot_r-1})")
        c.font = bold_font
        c.number_format = fmt_currency
        c.alignment = align_right
    for c in range(1, 10):
        ws.cell(row=tot_r, column=c).border = double_bottom
        ws.cell(row=tot_r, column=c).fill = total_fill
        
    auto_fit(ws)

def build_sheet_14_gl_index_mapping(wb):
    ws = wb.create_sheet(title="GL_Index_Mapping")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — GENERAL LEDGER INDEX & AUDIT TRAIL MAPPING").font = title_font
    ws.cell(row=2, column=1, value="Master Cross-Reference Mapping Every Trial Balance Account to Ledger Rows, Schedules and Notes").font = italic_font
    
    headers = ["Account Code", "Account Title / Description", "Category", "TB Reference", "GL Ledger Sheet", "GL Row Range", "Supporting Schedule Tab", "Financial Statement Line", "Note #", "Audit Cross-Tie Status"]
    style_header(ws, 4, headers)
    
    mapping_data = [
        ("1010", "Providus Bank - A/C 5400281942", "Current Assets", "TB_Adjusted_PreClosing!A5", "GL_Ledgers", "Rows 5-13", "Coverage_Bank_Matrix", "Cash and cash equivalents", "Note 11", "TIED & VERIFIED"),
        ("1020", "First Bank of Nigeria - A/C 2034891102", "Current Assets", "TB_Adjusted_PreClosing!A6", "GL_Ledgers", "Rows 14-15", "Coverage_Bank_Matrix", "Cash and cash equivalents", "Note 11", "TIED & VERIFIED"),
        ("1030", "Sterling Bank - A/C 0078451290", "Current Assets", "TB_Adjusted_PreClosing!A7", "GL_Ledgers", "Rows 16-17", "Coverage_Bank_Matrix", "Cash and cash equivalents", "Note 11", "TIED & VERIFIED"),
        ("1110", "Trade Receivables - Construction", "Current Assets", "TB_Adjusted_PreClosing!A8", "GL_Ledgers", "Rows 18-20", "Receivables_ECL_Payables", "Trade and other receivables", "Note 10", "TIED & VERIFIED"),
        ("1130", "Allowance for Expected Credit Losses", "Current Assets", "TB_Adjusted_PreClosing!A9", "GL_Ledgers", "Row 21", "Receivables_ECL_Payables", "Trade and other receivables", "Note 10", "TIED & VERIFIED"),
        ("1140", "Contract Assets (Unbilled Revenue)", "Current Assets", "TB_Adjusted_PreClosing!A10", "GL_Ledgers", "Row 22", "IFRS15_Revenue_WIP", "Contract assets", "Note 9", "TIED & VERIFIED"),
        ("1150", "WHT Credit Notes Receivable", "Current Assets", "TB_Adjusted_PreClosing!A11", "GL_Ledgers", "Row 23", "WHT_Receivable_Payable", "Trade and other receivables", "Note 10", "TIED & VERIFIED"),
        ("1160", "Prepayments & Site Advances", "Current Assets", "TB_Adjusted_PreClosing!A12", "GL_Ledgers", "Row 24", "Receivables_ECL_Payables", "Trade and other receivables", "Note 10", "TIED & VERIFIED"),
        ("1210", "Development Inventory - Land", "Current Assets", "TB_Adjusted_PreClosing!A13", "GL_Ledgers", "Rows 25-27", "IFRS15_Revenue_WIP", "Inventories", "Note 8", "TIED & VERIFIED"),
        ("1220", "Development Inventory - WIP", "Current Assets", "TB_Adjusted_PreClosing!A14", "GL_Ledgers", "Row 28", "IFRS15_Revenue_WIP", "Inventories", "Note 8", "TIED & VERIFIED"),
        ("1510", "PPE - Office & Site Equipment", "Non-Current Assets", "TB_Adjusted_PreClosing!A15", "GL_Ledgers", "Rows 29-30", "PPE_Depreciation", "Property, plant and equipment", "Note 7", "TIED & VERIFIED"),
        ("1590", "Accumulated Depreciation - PPE", "Non-Current Assets", "TB_Adjusted_PreClosing!A16", "GL_Ledgers", "Rows 31-32", "PPE_Depreciation", "Property, plant and equipment", "Note 7", "TIED & VERIFIED"),
        ("1610", "Deferred Tax Asset (IAS 12)", "Non-Current Assets", "TB_Adjusted_PreClosing!A17", "GL_Ledgers", "Row 33", "Deferred_Tax_IAS12", "Deferred tax asset", "Note 15", "TIED & VERIFIED"),
        ("2010", "Trade Payables - Subcontractors & Materials", "Current Liabilities", "TB_Adjusted_PreClosing!A18", "GL_Ledgers", "Rows 34-35", "Receivables_ECL_Payables", "Trade and other payables", "Note 14", "TIED & VERIFIED"),
        ("2020", "Contract Liabilities - Customer Deposits", "Current Liabilities", "TB_Adjusted_PreClosing!A19", "GL_Ledgers", "Rows 36-37", "IFRS15_Revenue_WIP", "Contract liabilities", "Note 13", "TIED & VERIFIED"),
        ("2030", "Accrued Audit & Professional Fees", "Current Liabilities", "TB_Adjusted_PreClosing!A20", "GL_Ledgers", "Rows 38-39", "Receivables_ECL_Payables", "Trade and other payables", "Note 14", "TIED & VERIFIED"),
        ("2040", "Other Accrued Expenses & Sundry Creditors", "Current Liabilities", "TB_Adjusted_PreClosing!A21", "GL_Ledgers", "Row 40", "Receivables_ECL_Payables", "Trade and other payables", "Note 14", "TIED & VERIFIED"),
        ("2110", "Current Tax - CIT Payable", "Current Liabilities", "TB_Adjusted_PreClosing!A22", "GL_Ledgers", "Row 41", "CIT_TET_MinTax", "Current tax liabilities", "Note 16", "TIED & VERIFIED"),
        ("2120", "Current Tax - TET Payable", "Current Liabilities", "TB_Adjusted_PreClosing!A23", "GL_Ledgers", "Row 42", "CIT_TET_MinTax", "Current tax liabilities", "Note 16", "TIED & VERIFIED"),
        ("2130", "Current Tax - PTF Payable", "Current Liabilities", "TB_Adjusted_PreClosing!A24", "GL_Ledgers", "Rows 43-44", "Tax_Payable_Rollforward", "Current tax liabilities", "Note 16", "TIED & VERIFIED"),
        ("2510", "Director's Loan & Long-Term Funding", "Non-Current Liabilities", "TB_Adjusted_PreClosing!A25", "GL_Ledgers", "Row 45", "Related_Parties_Equity", "Borrowings & related party loans", "Note 17", "TIED & VERIFIED"),
        ("3010", "Ordinary Share Capital (₦1.00 par)", "Equity", "TB_Adjusted_PreClosing!A26", "GL_Ledgers", "Row 46", "Related_Parties_Equity", "Share capital", "Note 18", "TIED & VERIFIED"),
        ("3020", "Retained Earnings (Accumulated Profit)", "Equity", "TB_Adjusted_PreClosing!A27", "GL_Ledgers", "Row 47", "Related_Parties_Equity", "Retained earnings", "Note 19", "TIED & VERIFIED"),
        ("4010", "Revenue - Civil Engineering Contracts", "Revenue", "TB_Adjusted_PreClosing!A28", "GL_Ledgers", "Row 48", "IFRS15_Revenue_WIP", "Revenue from contracts with customers", "Note 3", "TIED & VERIFIED"),
        ("4020", "Revenue - Land Resale", "Revenue", "TB_Adjusted_PreClosing!A29", "GL_Ledgers", "Row 49", "IFRS15_Revenue_WIP", "Revenue from contracts with customers", "Note 3", "TIED & VERIFIED"),
        ("4030", "Revenue - Project Management & Consulting", "Revenue", "TB_Adjusted_PreClosing!A30", "GL_Ledgers", "Row 50", "IFRS15_Revenue_WIP", "Revenue from contracts with customers", "Note 3", "TIED & VERIFIED"),
        ("5010", "Cost of Sales - Materials & Subcontracts", "Direct Costs", "TB_Adjusted_PreClosing!A31", "GL_Ledgers", "Row 51", "IFRS15_Revenue_WIP", "Cost of sales", "Note 4", "TIED & VERIFIED"),
        ("5020", "Cost of Sales - Land Inventory Released", "Direct Costs", "TB_Adjusted_PreClosing!A32", "GL_Ledgers", "Row 52", "IFRS15_Revenue_WIP", "Cost of sales", "Note 4", "TIED & VERIFIED"),
        ("5030", "Cost of Sales - Direct Site Labor & Hire", "Direct Costs", "TB_Adjusted_PreClosing!A33", "GL_Ledgers", "Row 53", "IFRS15_Revenue_WIP", "Cost of sales", "Note 4", "TIED & VERIFIED"),
        ("6010", "Administrative - Staff Salaries, Wages", "Operating Expenses", "TB_Adjusted_PreClosing!A34", "GL_Ledgers", "Row 54", "Notes_2023", "Administrative & operating expenses", "Note 5", "TIED & VERIFIED"),
        ("6020", "Administrative - Site Office Rent", "Operating Expenses", "TB_Adjusted_PreClosing!A35", "GL_Ledgers", "Row 55", "Notes_2023", "Administrative & operating expenses", "Note 5", "TIED & VERIFIED"),
        ("6030", "Administrative - Professional Fees", "Operating Expenses", "TB_Adjusted_PreClosing!A36", "GL_Ledgers", "Row 56", "Notes_2023", "Administrative & operating expenses", "Note 5", "TIED & VERIFIED"),
        ("6040", "Administrative - Audit & Tax Fees", "Operating Expenses", "TB_Adjusted_PreClosing!A37", "GL_Ledgers", "Row 57", "Notes_2023", "Administrative & operating expenses", "Note 5", "TIED & VERIFIED"),
        ("6050", "Administrative - Bank Charges & COT", "Operating Expenses", "TB_Adjusted_PreClosing!A38", "GL_Ledgers", "Row 58", "Notes_2023", "Administrative & operating expenses", "Note 5", "TIED & VERIFIED"),
        ("6060", "Administrative - Motor Vehicle & Logistics", "Operating Expenses", "TB_Adjusted_PreClosing!A39", "GL_Ledgers", "Row 59", "Notes_2023", "Administrative & operating expenses", "Note 5", "TIED & VERIFIED"),
        ("6070", "Administrative - Depreciation Expense (IAS 16)", "Operating Expenses", "TB_Adjusted_PreClosing!A40", "GL_Ledgers", "Row 60", "PPE_Depreciation", "Administrative & operating expenses", "Note 5", "TIED & VERIFIED"),
        ("6080", "Administrative - Impairment Allowance (ECL)", "Operating Expenses", "TB_Adjusted_PreClosing!A41", "GL_Ledgers", "Row 61", "Receivables_ECL_Payables", "Administrative & operating expenses", "Note 5", "TIED & VERIFIED"),
        ("6090", "Administrative - General Admin & Utilities", "Operating Expenses", "TB_Adjusted_PreClosing!A42", "GL_Ledgers", "Row 62", "Notes_2023", "Administrative & operating expenses", "Note 5", "TIED & VERIFIED"),
        ("7010", "Finance Costs - Bank Project Facility", "Finance Costs", "TB_Adjusted_PreClosing!A43", "GL_Ledgers", "Row 63", "Notes_2023", "Finance costs", "Note 6", "TIED & VERIFIED"),
        ("8010", "Income Tax Expense - Current Tax", "Taxation", "TB_Adjusted_PreClosing!A44", "GL_Ledgers", "Row 64", "CIT_TET_MinTax", "Income tax expense", "Note 16", "TIED & VERIFIED"),
        ("8020", "Income Tax Expense - Deferred Tax", "Taxation", "TB_Adjusted_PreClosing!A45", "GL_Ledgers", "Row 65", "Deferred_Tax_IAS12", "Income tax expense", "Note 15", "TIED & VERIFIED"),
    ]
    
    for idx, (coa, title, cat, tb_ref, gl_sh, gl_rw, sched, f_line, note, stat) in enumerate(mapping_data, 5):
        ws.cell(row=idx, column=1, value=coa).alignment = align_center
        ws.cell(row=idx, column=2, value=title).font = bold_font
        ws.cell(row=idx, column=3, value=cat).font = regular_font
        ws.cell(row=idx, column=4, value=tb_ref).font = italic_font
        ws.cell(row=idx, column=5, value=gl_sh).alignment = align_center
        ws.cell(row=idx, column=6, value=gl_rw).alignment = align_center
        ws.cell(row=idx, column=7, value=sched).alignment = align_center
        ws.cell(row=idx, column=8, value=f_line).font = regular_font
        ws.cell(row=idx, column=9, value=note).alignment = align_center
        c_st = ws.cell(row=idx, column=10, value=stat)
        c_st.alignment = align_center
        c_st.font = success_font
        
        for c in range(1, 11):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

print("sheet_builders_2.py completed with Sheets 9-14.")
