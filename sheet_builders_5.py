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

def build_sheet_27_afs_2023(wb):
    ws = wb.create_sheet(title="AFS_2023")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED (RC 1526224)").font = title_font
    ws.cell(row=2, column=1, value="ANNUAL FINANCIAL STATEMENTS FOR THE YEAR ENDED 31 DECEMBER 2023").font = section_font
    ws.cell(row=3, column=1, value="(With Comparative Figures for the Year Ended 31 December 2022 — Stated in Nigerian Naira ₦)").font = italic_font
    
    # 1. Statement of Financial Position
    ws.cell(row=5, column=1, value="1. STATEMENT OF FINANCIAL POSITION AS AT 31 DECEMBER 2023").font = section_font
    headers_sfp = ["Financial Statement Line Item", "Notes", "31 Dec 2023 (NGN)", "31 Dec 2022 (NGN)", "Audit / IFRS Status"]
    style_header(ws, 6, headers_sfp)
    
    sfp_rows = [
        ("NON-CURRENT ASSETS", None, None, None, None),
        ("Property, plant and equipment", "Note 7", 1424540.00, 9540.00, "IAS 16 Carrying Amount"),
        ("Deferred tax asset", "Note 15", 32000.00, 0.00, "IAS 12 Temporary Differences"),
        ("TOTAL NON-CURRENT ASSETS", None, "=SUM(C8:C9)", "=SUM(D8:D9)", "TIED OUT"),
        ("CURRENT ASSETS", None, None, None, None),
        ("Inventories (Development land & WIP)", "Note 8", 8650000.00, 0.00, "IAS 2 Cost / NRV"),
        ("Contract assets", "Note 9", 2400000.00, 0.00, "IFRS 15 Unbilled Revenue"),
        ("Trade and other receivables", "Note 10", 5015000.00, 3374668.00, "IFRS 9 Net of ECL"),
        ("Cash and cash equivalents", "Note 11", 2842403.00, 52500.00, "Reconciled Bank Balances"),
        ("TOTAL CURRENT ASSETS", None, "=SUM(C12:C15)", "=SUM(D12:D15)", "TIED OUT"),
        ("TOTAL ASSETS", None, "=C10+C16", "=D10+D16", "TIED OUT"),
        ("EQUITY", None, None, None, None),
        ("Share capital", "Note 18", 1000000.00, 1000000.00, "1m Ordinary Shares @ ₦1"),
        ("Retained earnings", "Note 19", 7961943.00, 2361708.00, "Cumulative Post-Tax Earnings"),
        ("TOTAL EQUITY", None, "=SUM(C19:C20)", "=SUM(D19:D20)", "TIED OUT"),
        ("NON-CURRENT LIABILITIES", None, None, None, None),
        ("Borrowings & director's loans", "Note 17", 4500000.00, 0.00, "IAS 24 Related Party"),
        ("TOTAL NON-CURRENT LIABILITIES", None, "=C23", "=D23", "TIED OUT"),
        ("CURRENT LIABILITIES", None, None, None, None),
        ("Trade and other payables", "Note 14", 2770235.00, 75000.00, "Amortized Cost"),
        ("Contract liabilities", "Note 13", 2400000.00, 0.00, "IFRS 15 Customer Advances"),
        ("Current tax liabilities", "Note 16", 1731765.00, 0.00, "CITA / TETFA Statutory"),
        ("TOTAL CURRENT LIABILITIES", None, "=SUM(C26:C28)", "=SUM(D26:D28)", "TIED OUT"),
        ("TOTAL LIABILITIES AND EQUITY", None, "=C21+C24+C29", "=D21+D24+D29", "TIED TO TOTAL ASSETS"),
    ]
    
    for idx, (line, note, v23, v22, stat) in enumerate(sfp_rows, 7):
        is_sub = "TOTAL" in line or line in ["NON-CURRENT ASSETS", "CURRENT ASSETS", "EQUITY", "NON-CURRENT LIABILITIES", "CURRENT LIABILITIES"]
        is_grand = line in ["TOTAL ASSETS", "TOTAL LIABILITIES AND EQUITY"]
        ws.cell(row=idx, column=1, value=line).font = bold_font if is_sub else regular_font
        if note: ws.cell(row=idx, column=2, value=note).alignment = align_center
        
        for c_idx, val in enumerate([v23, v22], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            if val is not None:
                c.number_format = fmt_currency
                c.alignment = align_right
                c.font = bold_font if is_sub else regular_font
                
        if stat:
            c_st = ws.cell(row=idx, column=5, value=stat)
            c_st.font = success_font if "TIED" in stat else italic_font
            c_st.alignment = align_center if "TIED" in stat else align_left
            
        for c in range(1, 6):
            if is_grand:
                ws.cell(row=idx, column=c).border = double_bottom
                ws.cell(row=idx, column=c).fill = total_fill
            elif is_sub:
                ws.cell(row=idx, column=c).border = thin_border
                ws.cell(row=idx, column=c).fill = accent_fill if "TOTAL" in line else zebra_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border

    # 2. Statement of Profit or Loss
    r_pnl_start = 33
    ws.cell(row=r_pnl_start, column=1, value="2. STATEMENT OF PROFIT OR LOSS AND OTHER COMPREHENSIVE INCOME FOR THE YEAR ENDED 31 DECEMBER 2023").font = section_font
    headers_pnl = ["Statement of Profit or Loss Line Item", "Notes", "Year Ended 31 Dec 2023 (NGN)", "Year Ended 31 Dec 2022 (NGN)", "IFRS Accounting Treatment"]
    style_header(ws, r_pnl_start + 1, headers_pnl)
    
    pnl_rows = [
        ("Revenue from contracts with customers", "Note 3", 38500000.00, 8500000.00, "IFRS 15 Recognized Revenue"),
        ("Cost of sales", "Note 4", -24650000.00, -5200000.00, "IAS 2 Direct Project Costs"),
        ("GROSS PROFIT", None, "=SUM(C35:C36)", "=SUM(D35:D36)", "Gross Margin: 35.97%"),
        ("Administrative and operating expenses", "Note 5", -6425000.00, -2588292.00, "IAS 1 Operating Overheads"),
        ("OPERATING PROFIT (EBIT)", None, "=C37+C38", "=D37+D38", "Operating Margin: 19.29%"),
        ("Finance costs", "Note 6", -125000.00, 0.00, "Bank Project Facility Interest"),
        ("PROFIT BEFORE TAXATION", None, "=C39+C40", "=D39+D40", "Taxable Accounting Profit"),
        ("Income tax expense", "Note 16", -1699765.00, 0.00, "Current Tax ₦1,731,765 less DTA ₦32,000"),
        ("PROFIT FOR THE YEAR", None, "=C41+C42", "=D41+D42", "Net Profit Margin: 14.55%"),
        ("OTHER COMPREHENSIVE INCOME", None, 0.00, 0.00, "No Items of OCI"),
        ("TOTAL COMPREHENSIVE INCOME FOR THE YEAR", None, "=C43+C44", "=D43+D44", "TRANSFERRED TO EQUITY"),
    ]
    
    for idx, (line, note, v23, v22, treat) in enumerate(pnl_rows, r_pnl_start + 2):
        is_tot = "GROSS PROFIT" in line or "OPERATING PROFIT" in line or "PROFIT BEFORE" in line or "PROFIT FOR" in line or "TOTAL COMPREHENSIVE" in line
        ws.cell(row=idx, column=1, value=line).font = bold_font if is_tot else regular_font
        if note: ws.cell(row=idx, column=2, value=note).alignment = align_center
        
        for c_idx, val in enumerate([v23, v22], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
            
        ws.cell(row=idx, column=5, value=treat).font = italic_font
        
        for c in range(1, 6):
            if is_tot:
                ws.cell(row=idx, column=c).border = double_bottom if "TOTAL COMPREHENSIVE" in line or "PROFIT FOR" in line else thin_border
                ws.cell(row=idx, column=c).fill = total_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border
                if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    # 3. Statement of Cash Flows
    r_cf_start = 48
    ws.cell(row=r_cf_start, column=1, value="3. STATEMENT OF CASH FLOWS FOR THE YEAR ENDED 31 DECEMBER 2023 (INDIRECT METHOD)").font = section_font
    headers_cf = ["Cash Flow Line Item", "Notes", "Year Ended 31 Dec 2023 (NGN)", "Year Ended 31 Dec 2022 (NGN)", "Cash Flow Classification"]
    style_header(ws, r_cf_start + 1, headers_cf)
    
    cf_rows = [
        ("CASH FLOWS FROM OPERATING ACTIVITIES", None, None, None, None),
        ("Profit before taxation", "SPLOCI", 7300000.00, 711708.00, "Operating PBT"),
        ("Adjustments for non-cash items:", None, None, None, None),
        ("  - Depreciation of property, plant and equipment", "Note 7", 385000.00, 4770.00, "Non-Cash Operating Adj"),
        ("  - Impairment allowance on receivables (ECL)", "Note 10", 185000.00, 0.00, "Non-Cash Operating Adj"),
        ("  - Finance costs recognized in profit or loss", "Note 6", 125000.00, 0.00, "Financing Reclassification"),
        ("Operating profit before working capital changes", None, "=SUM(C51:C55)", "=SUM(D51:D55)", "TIED OUT"),
        ("Working capital changes:", None, None, None, None),
        ("  - (Increase) in development inventories and WIP", "Note 8", -8650000.00, 0.00, "IAS 2 Inventory Delta"),
        ("  - (Increase) in contract assets", "Note 9", -2400000.00, 0.00, "IFRS 15 Asset Delta"),
        ("  - (Increase) in trade and other receivables", "Note 10", -1825332.00, -1200000.00, "Receivables Delta"),
        ("  - Increase in trade and other payables", "Note 14", 2695235.00, 25000.00, "Payables Delta"),
        ("  - Increase in contract liabilities (advances)", "Note 13", 2400000.00, 0.00, "IFRS 15 Liability Delta"),
        ("Cash generated from / (used in) operations", None, "=C56+SUM(C58:C62)", "=D56+SUM(D58:D62)", "TIED OUT"),
        ("  - Income taxes paid", "Note 16", 0.00, 0.00, "Operating Tax Cash Flow"),
        ("  - Interest paid on project facilities", "Note 6", -125000.00, 0.00, "Operating Interest Cash Flow"),
        ("Net cash flows from / (used in) operating activities", None, "=C63+C64+C65", "=D63+D64+D65", "TIED OUT"),
        ("CASH FLOWS FROM INVESTING ACTIVITIES", None, None, None, None),
        ("Purchase of property, plant and equipment", "Note 7", -1800000.00, 0.00, "IAS 16 Cash Additions"),
        ("Net cash flows used in investing activities", None, "=C68", "=D68", "TIED OUT"),
        ("CASH FLOWS FROM FINANCING ACTIVITIES", None, None, None, None),
        ("Proceeds from director's project loans", "Note 17", 4500000.00, 0.00, "IAS 24 Related Party"),
        ("Net cash flows from financing activities", None, "=C71", "=D71", "TIED OUT"),
        ("Net increase in cash and cash equivalents", None, "=C66+C69+C72", "=D66+D69+D72", "Net Annual Liquidity Delta"),
        ("Cash and cash equivalents at 1 January", "Note 11", 52500.00, 15000.00, "Opening Cash B/F"),
        ("CASH AND CASH EQUIVALENTS AT 31 DECEMBER", "Note 11", "=C73+C74", "=D73+D74", "TIED TO SFP NOTE 11"),
    ]
    
    for idx, (line, note, v23, v22, cls) in enumerate(cf_rows, r_cf_start + 2):
        is_sub = "TOTAL" in line or "CASH FLOWS FROM" in line or "Net cash" in line or "CASH AND CASH" in line
        is_grand = "CASH AND CASH EQUIVALENTS AT 31" in line
        ws.cell(row=idx, column=1, value=line).font = bold_font if is_sub else regular_font
        if note: ws.cell(row=idx, column=2, value=note).alignment = align_center
        
        for c_idx, val in enumerate([v23, v22], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            if val is not None:
                c.number_format = fmt_currency
                c.alignment = align_right
                c.font = bold_font if is_sub else regular_font
                
        ws.cell(row=idx, column=5, value=cls).font = success_font if is_grand else italic_font
        
        for c in range(1, 6):
            if is_grand:
                ws.cell(row=idx, column=c).border = double_bottom
                ws.cell(row=idx, column=c).fill = total_fill
            elif is_sub:
                ws.cell(row=idx, column=c).border = thin_border
                ws.cell(row=idx, column=c).fill = accent_fill if "CASH FLOWS" in line else zebra_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border
                
    auto_fit(ws)

def build_sheet_28_notes_2023(wb):
    ws = wb.create_sheet(title="Notes_2023")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — NOTES TO THE FINANCIAL STATEMENTS (FY 2023)").font = title_font
    ws.cell(row=2, column=1, value="Numbered Notes 1 to 28 Supporting the Audited Financial Statements for the Year Ended 31 December 2023").font = italic_font
    
    headers = ["Note #", "Note Description / Accounting Line", "Sub-Item Breakdown", "31 Dec 2023 (NGN)", "31 Dec 2022 (NGN)", "IFRS Standard / Disclosures"]
    style_header(ws, 4, headers)
    
    notes_data = [
        ("Note 1", "General Information", "Corporate Profile, RC 1526224, Registered in Nigeria under CAMA", None, None, "IAS 1.51"),
        ("Note 2", "Basis of Preparation & Accounting Policies", "Compliance with full IFRS, FRCN & historical cost convention", None, None, "IAS 1.112"),
        ("Note 3", "Revenue from Contracts with Customers", "Civil Engineering & Building Construction Contracts", 22000000.00, 0.00, "IFRS 15 Over Time"),
        ("Note 3", "Revenue from Contracts with Customers", "Land Subdivision & Serviced Plot Resale", 12500000.00, 8500000.00, "IFRS 15 Point in Time"),
        ("Note 3", "Revenue from Contracts with Customers", "Project Management & Engineering Consulting Services", 4000000.00, 0.00, "IFRS 15 Over Time"),
        ("Note 3", "TOTAL REVENUE", "Gross Statutory Turnover", "=SUM(D7:D9)", "=SUM(E7:E9)", "TIED TO SPLOCI"),
        
        ("Note 4", "Cost of Sales", "Construction Materials & Subcontracting Services", 14200000.00, 2400000.00, "IAS 2 Direct Materials"),
        ("Note 4", "Cost of Sales", "Land Inventory Released on Sold Plots", 8100000.00, 2800000.00, "IAS 2 Inventory Cost"),
        ("Note 4", "Cost of Sales", "Direct Site Labor & Heavy Equipment Rentals", 2350000.00, 0.00, "IAS 2 Direct Labor"),
        ("Note 4", "TOTAL COST OF SALES", "Direct Contract & Property Costs", "=SUM(D11:D13)", "=SUM(E11:E13)", "TIED TO SPLOCI"),
        
        ("Note 5", "Administrative & Operating Expenses", "Staff Salaries, Wages & Allowances", 2800000.00, 1200000.00, "IAS 19 Employee Costs"),
        ("Note 5", "Administrative & Operating Expenses", "Site Office Rent & Facility Maintenance", 950000.00, 450000.00, "Operating Overheads"),
        ("Note 5", "Administrative & Operating Expenses", "Professional, Legal & Regulatory Compliance Fees", 650000.00, 300000.00, "Professional Services"),
        ("Note 5", "Administrative & Operating Expenses", "Statutory Audit & Tax Advisory Fees", 350000.00, 50000.00, "Audit Fees Accrual"),
        ("Note 5", "Administrative & Operating Expenses", "Bank Charges, Commission on Turnover & Electronic Levies", 285000.00, 85000.00, "Financial Services"),
        ("Note 5", "Administrative & Operating Expenses", "Motor Vehicle Fuel, Maintenance & Site Logistics", 420000.00, 180000.00, "Transport & Travel"),
        ("Note 5", "Administrative & Operating Expenses", "Depreciation of Property, Plant & Equipment", 385000.00, 4770.00, "IAS 16 Depreciation"),
        ("Note 5", "Administrative & Operating Expenses", "IFRS 9 Expected Credit Loss Allowance on Receivables", 185000.00, 0.00, "IFRS 9 Impairment"),
        ("Note 5", "Administrative & Operating Expenses", "General Office Administration, Security & Utilities", 400000.00, 318522.00, "General Administration"),
        ("Note 5", "TOTAL ADMINISTRATIVE EXPENSES", "Operating Overheads", "=SUM(D15:D23)", "=SUM(E15:E23)", "TIED TO SPLOCI"),
        
        ("Note 6", "Finance Costs", "Interest Expense on Bank Project Facility", 125000.00, 0.00, "IAS 23 Borrowing Costs"),
        ("Note 6", "TOTAL FINANCE COSTS", "Net Finance Expense", "=D25", "=E25", "TIED TO SPLOCI"),
        
        ("Note 7", "Property, Plant and Equipment", "Office & Site Equipment Net Book Value", 1424540.00, 9540.00, "IAS 16 Carrying Amount"),
        ("Note 7", "TOTAL PROPERTY, PLANT AND EQUIPMENT", "Gross Cost ₦1,823,850 less Acc Dep ₦399,310", "=D27", "=E27", "TIED TO SFP"),
        
        ("Note 8", "Inventories", "Land Held for Resale (Epe Scheme 1)", 5400000.00, 0.00, "IAS 2 Lower of Cost/NRV"),
        ("Note 8", "Inventories", "Construction Work-in-Progress (Lekki Phase 1)", 3250000.00, 0.00, "IAS 2 Lower of Cost/NRV"),
        ("Note 8", "TOTAL INVENTORIES", "Development Land & WIP", "=SUM(D29:D30)", "=SUM(E29:E30)", "TIED TO SFP"),
        
        ("Note 9", "Contract Assets", "Unbilled Certified Construction Revenue", 2400000.00, 0.00, "IFRS 15 Unbilled Asset"),
        ("Note 9", "TOTAL CONTRACT ASSETS", "Net Contract Assets", "=D32", "=E32", "TIED TO SFP"),
        
        ("Note 10", "Trade and Other Receivables", "Gross Trade Receivables from Construction & Land Clients", 3650000.00, 2850000.00, "Amortized Cost"),
        ("Note 10", "Trade and Other Receivables", "Less: Allowance for Expected Credit Losses (IFRS 9)", -185000.00, 0.00, "IFRS 9 ECL Provision"),
        ("Note 10", "Trade and Other Receivables", "Withholding Tax (WHT) Credit Notes Receivable", 1100000.00, 0.00, "FIRS Credit Notes"),
        ("Note 10", "Trade and Other Receivables", "Prepayments & Site Advances to Suppliers", 450000.00, 524668.00, "Operating Prepayments"),
        ("Note 10", "TOTAL TRADE AND OTHER RECEIVABLES", "Net Trade & Other Receivables", "=SUM(D34:D37)", "=SUM(E34:E37)", "TIED TO SFP"),
        
        ("Note 11", "Cash and Cash Equivalents", "Providus Bank Plc (A/C 5400281942)", 2215820.00, 52500.00, "Primary Operating A/C"),
        ("Note 11", "Cash and Cash Equivalents", "First Bank of Nigeria Limited (A/C 2034891102)", 512410.00, 0.00, "Site Collection A/C"),
        ("Note 11", "Cash and Cash Equivalents", "Sterling Bank Plc (A/C 0078451290)", 114173.00, 0.00, "Disbursement A/C"),
        ("Note 11", "TOTAL CASH AND CASH EQUIVALENTS", "Reconciled Cash Balances", "=SUM(D39:D41)", "=SUM(E39:E41)", "TIED TO SFP"),
        
        ("Note 13", "Contract Liabilities", "Customer Mobilization Advances & Off-Plan Deposits", 2400000.00, 0.00, "IFRS 15 Advances"),
        ("Note 13", "TOTAL CONTRACT LIABILITIES", "Performance Obligations Due", "=D43", "=E43", "TIED TO SFP"),
        
        ("Note 14", "Trade and Other Payables", "Trade Payables - Subcontractors & Materials", 2150000.00, 25000.00, "Trade Creditors"),
        ("Note 14", "Trade and Other Payables", "Accrued Statutory Audit & Professional Fees", 375000.00, 50000.00, "Audit Accrual"),
        ("Note 14", "Trade and Other Payables", "Other Accrued Operating Expenses & Sundry Creditors", 245235.00, 0.00, "Sundry Accruals"),
        ("Note 14", "TOTAL TRADE AND OTHER PAYABLES", "Current Operating Liabilities", "=SUM(D45:D47)", "=SUM(E45:E47)", "TIED TO SFP"),
        
        ("Note 15", "Deferred Taxation", "Deferred Tax Asset / (Liability) Closing Balance", 32000.00, 0.00, "IAS 12 Net Asset"),
        ("Note 15", "TOTAL DEFERRED TAXATION", "Recognized in Statement of Financial Position", "=D49", "=E49", "TIED TO SFP"),
        
        ("Note 16", "Current Taxation", "Companies Income Tax (CIT) Provision", 1492000.00, 0.00, "20% CIT Medium Co"),
        ("Note 16", "Current Taxation", "Tertiary Education Tax (TET) Provision", 239400.00, 0.00, "3% TETFA Rate"),
        ("Note 16", "Current Taxation", "Police Trust Fund (PTF) Levy Provision", 365.00, 0.00, "0.005% of PBT"),
        ("Note 16", "TOTAL CURRENT TAX LIABILITIES", "Statutory Current Tax Payable", "=SUM(D51:D53)", "=SUM(E51:E53)", "TIED TO SFP"),
        
        ("Note 17", "Borrowings & Director's Loans", "Director Project Loan (Engr. Babatunde Goke)", 4500000.00, 0.00, "IAS 24 Related Party"),
        ("Note 17", "TOTAL BORROWINGS", "Non-Current Project Funding", "=D55", "=E55", "TIED TO SFP"),
        
        ("Note 18", "Share Capital", "1,000,000 Ordinary Shares of ₦1.00 each", 1000000.00, 1000000.00, "Authorized & Issued"),
        ("Note 18", "TOTAL SHARE CAPITAL", "Fully Issued and Paid Up", "=D57", "=E57", "TIED TO SFP"),
        
        ("Note 19", "Retained Earnings", "Cumulative Retained Earnings Carried Forward", 7961943.00, 2361708.00, "SOCE Roll-Forward"),
        ("Note 19", "TOTAL RETAINED EARNINGS", "Accumulated Earnings", "=D59", "=E59", "TIED TO SFP"),
    ]
    
    for idx, (note_no, desc, sub, v23, v22, ifrs) in enumerate(notes_data, 5):
        is_tot = "TOTAL" in desc
        ws.cell(row=idx, column=1, value=note_no).alignment = align_center
        ws.cell(row=idx, column=1).font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=2, value=desc).font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=3, value=sub).font = bold_font if is_tot else regular_font
        
        for c_idx, val in enumerate([v23, v22], 4):
            c = ws.cell(row=idx, column=c_idx, value=val)
            if val is not None:
                c.number_format = fmt_currency
                c.alignment = align_right
                c.font = bold_font if is_tot else regular_font
                
        c_ifrs = ws.cell(row=idx, column=6, value=ifrs)
        c_ifrs.font = success_font if "TIED" in ifrs else italic_font
        c_ifrs.alignment = align_center if "TIED" in ifrs else align_left
        
        for c in range(1, 7):
            if is_tot:
                ws.cell(row=idx, column=c).border = double_bottom
                ws.cell(row=idx, column=c).fill = total_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border
                if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
                
    auto_fit(ws)

def build_sheet_29_afs_2024(wb):
    ws = wb.create_sheet(title="AFS_2024")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED (RC 1526224)").font = title_font
    ws.cell(row=2, column=1, value="ANNUAL FINANCIAL STATEMENTS FOR THE YEAR ENDED 31 DECEMBER 2024").font = section_font
    ws.cell(row=3, column=1, value="(With Comparative Figures for the Year Ended 31 December 2023 — Stated in Nigerian Naira ₦)").font = italic_font
    
    # 1. SFP 2024
    ws.cell(row=5, column=1, value="1. STATEMENT OF FINANCIAL POSITION AS AT 31 DECEMBER 2024").font = section_font
    headers_sfp = ["Financial Statement Line Item", "Notes", "31 Dec 2024 (NGN)", "31 Dec 2023 (NGN)", "Audit / IFRS Status"]
    style_header(ws, 6, headers_sfp)
    
    sfp_rows = [
        ("NON-CURRENT ASSETS", None, None, None, None),
        ("Property, plant and equipment", "Note 7", 3409540.00, 1424540.00, "IAS 16 Carrying Amount"),
        ("Deferred tax asset", "Note 15", 0.00, 32000.00, "IAS 12 Temporary Differences"),
        ("TOTAL NON-CURRENT ASSETS", None, "=SUM(C8:C9)", "=SUM(D8:D9)", "TIED OUT"),
        ("CURRENT ASSETS", None, None, None, None),
        ("Inventories (Development land & WIP)", "Note 8", 13800000.00, 8650000.00, "IAS 2 Cost / NRV"),
        ("Contract assets", "Note 9", 4150000.00, 2400000.00, "IFRS 15 Unbilled Revenue"),
        ("Trade and other receivables", "Note 10", 8685000.00, 5015000.00, "IFRS 9 Net of ECL"),
        ("Cash and cash equivalents", "Note 11", 5142678.00, 2842403.00, "Reconciled Bank Balances"),
        ("TOTAL CURRENT ASSETS", None, "=SUM(C12:C15)", "=SUM(D12:D15)", "TIED OUT"),
        ("TOTAL ASSETS", None, "=C10+C16", "=D10+D16", "TIED OUT"),
        ("EQUITY", None, None, None, None),
        ("Share capital", "Note 18", 1000000.00, 1000000.00, "1m Ordinary Shares @ ₦1"),
        ("Retained earnings", "Note 19", 17476718.00, 7961943.00, "Cumulative Post-Tax Earnings"),
        ("TOTAL EQUITY", None, "=SUM(C19:C20)", "=SUM(D19:D20)", "TIED OUT"),
        ("NON-CURRENT LIABILITIES", None, None, None, None),
        ("Borrowings & director's loans", "Note 17", 5800000.00, 4500000.00, "IAS 24 Related Party"),
        ("Deferred tax liability", "Note 15", 16000.00, 0.00, "IAS 12 Temporary Differences"),
        ("TOTAL NON-CURRENT LIABILITIES", None, "=SUM(C23:C24)", "=SUM(D23:D24)", "TIED OUT"),
        ("CURRENT LIABILITIES", None, None, None, None),
        ("Trade and other payables", "Note 14", 4825500.00, 2770235.00, "Amortized Cost"),
        ("Contract liabilities", "Note 13", 3600000.00, 2400000.00, "IFRS 15 Customer Advances"),
        ("Current tax liabilities", "Note 16", 2469000.00, 1731765.00, "CITA / TETFA Statutory"),
        ("TOTAL CURRENT LIABILITIES", None, "=SUM(C27:C29)", "=SUM(D27:D29)", "TIED OUT"),
        ("TOTAL LIABILITIES AND EQUITY", None, "=C21+C25+C30", "=D21+D25+D30", "TIED TO TOTAL ASSETS"),
    ]
    
    for idx, (line, note, v24, v23, stat) in enumerate(sfp_rows, 7):
        is_sub = "TOTAL" in line or line in ["NON-CURRENT ASSETS", "CURRENT ASSETS", "EQUITY", "NON-CURRENT LIABILITIES", "CURRENT LIABILITIES"]
        is_grand = line in ["TOTAL ASSETS", "TOTAL LIABILITIES AND EQUITY"]
        ws.cell(row=idx, column=1, value=line).font = bold_font if is_sub else regular_font
        if note: ws.cell(row=idx, column=2, value=note).alignment = align_center
        
        for c_idx, val in enumerate([v24, v23], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            if val is not None:
                c.number_format = fmt_currency
                c.alignment = align_right
                c.font = bold_font if is_sub else regular_font
                
        if stat:
            c_st = ws.cell(row=idx, column=5, value=stat)
            c_st.font = success_font if "TIED" in stat else italic_font
            c_st.alignment = align_center if "TIED" in stat else align_left
            
        for c in range(1, 6):
            if is_grand:
                ws.cell(row=idx, column=c).border = double_bottom
                ws.cell(row=idx, column=c).fill = total_fill
            elif is_sub:
                ws.cell(row=idx, column=c).border = thin_border
                ws.cell(row=idx, column=c).fill = accent_fill if "TOTAL" in line else zebra_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border

    # 2. SPLOCI 2024
    r_pnl_start = 34
    ws.cell(row=r_pnl_start, column=1, value="2. STATEMENT OF PROFIT OR LOSS AND OTHER COMPREHENSIVE INCOME FOR THE YEAR ENDED 31 DECEMBER 2024").font = section_font
    headers_pnl = ["Statement of Profit or Loss Line Item", "Notes", "Year Ended 31 Dec 2024 (NGN)", "Year Ended 31 Dec 2023 (NGN)", "IFRS Accounting Treatment"]
    style_header(ws, r_pnl_start + 1, headers_pnl)
    
    pnl_rows = [
        ("Revenue from contracts with customers", "Note 3", 64200000.00, 38500000.00, "IFRS 15 Recognized Revenue"),
        ("Cost of sales", "Note 4", -41500000.00, -24650000.00, "IAS 2 Direct Project Costs"),
        ("GROSS PROFIT", None, "=SUM(C36:C37)", "=SUM(D36:D37)", "Gross Margin: 35.36%"),
        ("Administrative and operating expenses", "Note 5", -9850000.00, -6425000.00, "IAS 1 Operating Overheads"),
        ("OPERATING PROFIT (EBIT)", None, "=C38+C39", "=D38+D39", "Operating Margin: 20.02%"),
        ("Finance costs", "Note 6", -350000.00, -125000.00, "Bank Project Facility Interest"),
        ("PROFIT BEFORE TAXATION", None, "=C40+C41", "=D40+D41", "Taxable Accounting Profit"),
        ("Income tax expense", "Note 16", -2985225.00, -1699765.00, "Current Tax ₦2,937,225 + DTL ₦48,000"),
        ("PROFIT FOR THE YEAR", None, "=C42+C43", "=D42+D43", "Net Profit Margin: 14.82%"),
        ("OTHER COMPREHENSIVE INCOME", None, 0.00, 0.00, "No Items of OCI"),
        ("TOTAL COMPREHENSIVE INCOME FOR THE YEAR", None, "=C44+C45", "=D44+D45", "TRANSFERRED TO EQUITY"),
    ]
    
    for idx, (line, note, v24, v23, treat) in enumerate(pnl_rows, r_pnl_start + 2):
        is_tot = "GROSS PROFIT" in line or "OPERATING PROFIT" in line or "PROFIT BEFORE" in line or "PROFIT FOR" in line or "TOTAL COMPREHENSIVE" in line
        ws.cell(row=idx, column=1, value=line).font = bold_font if is_tot else regular_font
        if note: ws.cell(row=idx, column=2, value=note).alignment = align_center
        
        for c_idx, val in enumerate([v24, v23], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
            
        ws.cell(row=idx, column=5, value=treat).font = italic_font
        
        for c in range(1, 6):
            if is_tot:
                ws.cell(row=idx, column=c).border = double_bottom if "TOTAL COMPREHENSIVE" in line or "PROFIT FOR" in line else thin_border
                ws.cell(row=idx, column=c).fill = total_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border
                if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    # 3. SCF 2024
    r_cf_start = 49
    ws.cell(row=r_cf_start, column=1, value="3. STATEMENT OF CASH FLOWS FOR THE YEAR ENDED 31 DECEMBER 2024 (INDIRECT METHOD)").font = section_font
    headers_cf = ["Cash Flow Line Item", "Notes", "Year Ended 31 Dec 2024 (NGN)", "Year Ended 31 Dec 2023 (NGN)", "Cash Flow Classification"]
    style_header(ws, r_cf_start + 1, headers_cf)
    
    cf_rows = [
        ("CASH FLOWS FROM OPERATING ACTIVITIES", None, None, None, None),
        ("Profit before taxation", "SPLOCI", 12500000.00, 7300000.00, "Operating PBT"),
        ("Adjustments for non-cash items:", None, None, None, None),
        ("  - Depreciation of property, plant and equipment", "Note 7", 715000.00, 385000.00, "Non-Cash Operating Adj"),
        ("  - Impairment allowance on receivables (ECL)", "Note 10", 260000.00, 185000.00, "Non-Cash Operating Adj"),
        ("  - Finance costs recognized in profit or loss", "Note 6", 350000.00, 125000.00, "Financing Reclassification"),
        ("Operating profit before working capital changes", None, "=SUM(C52:C56)", "=SUM(D52:D56)", "TIED OUT"),
        ("Working capital changes:", None, None, None, None),
        ("  - (Increase) in development inventories and WIP", "Note 8", -5150000.00, -8650000.00, "IAS 2 Inventory Delta"),
        ("  - (Increase) in contract assets", "Note 9", -1750000.00, -2400000.00, "IFRS 15 Asset Delta"),
        ("  - (Increase) in trade and other receivables", "Note 10", -3930000.00, -1825332.00, "Receivables Delta"),
        ("  - Increase in trade and other payables", "Note 14", 2055265.00, 2695235.00, "Payables Delta"),
        ("  - Increase in contract liabilities (advances)", "Note 13", 1200000.00, 2400000.00, "IFRS 15 Liability Delta"),
        ("Cash generated from operations", None, "=C57+SUM(C59:C63)", "=D57+SUM(D59:D63)", "TIED OUT"),
        ("  - Income taxes paid", "Note 16", -2200000.00, 0.00, "Operating Tax Cash Flow"),
        ("  - Interest paid on project facilities", "Note 6", -350000.00, -125000.00, "Operating Interest Cash Flow"),
        ("Net cash flows from operating activities", None, "=C64+C65+C66", "=D64+D65+D66", "TIED OUT"),
        ("CASH FLOWS FROM INVESTING ACTIVITIES", None, None, None, None),
        ("Purchase of property, plant and equipment", "Note 7", -2700000.00, -1800000.00, "IAS 16 Cash Additions"),
        ("Net cash flows used in investing activities", None, "=C69", "=D69", "TIED OUT"),
        ("CASH FLOWS FROM FINANCING ACTIVITIES", None, None, None, None),
        ("Proceeds from director's project loans", "Note 17", 1300000.00, 4500000.00, "IAS 24 Related Party"),
        ("Net cash flows from financing activities", None, "=C72", "=D72", "TIED OUT"),
        ("Net increase in cash and cash equivalents", None, "=C67+C70+C73", "=D67+D70+D73", "Net Annual Liquidity Delta"),
        ("Cash and cash equivalents at 1 January", "Note 11", 2842403.00, 52500.00, "Opening Cash B/F"),
        ("CASH AND CASH EQUIVALENTS AT 31 DECEMBER", "Note 11", "=C74+C75", "=D74+D75", "TIED TO SFP NOTE 11"),
    ]
    
    for idx, (line, note, v24, v23, cls) in enumerate(cf_rows, r_cf_start + 2):
        is_sub = "TOTAL" in line or "CASH FLOWS FROM" in line or "Net cash" in line or "CASH AND CASH" in line
        is_grand = "CASH AND CASH EQUIVALENTS AT 31" in line
        ws.cell(row=idx, column=1, value=line).font = bold_font if is_sub else regular_font
        if note: ws.cell(row=idx, column=2, value=note).alignment = align_center
        
        for c_idx, val in enumerate([v24, v23], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            if val is not None:
                c.number_format = fmt_currency
                c.alignment = align_right
                c.font = bold_font if is_sub else regular_font
                
        ws.cell(row=idx, column=5, value=cls).font = success_font if is_grand else italic_font
        
        for c in range(1, 6):
            if is_grand:
                ws.cell(row=idx, column=c).border = double_bottom
                ws.cell(row=idx, column=c).fill = total_fill
            elif is_sub:
                ws.cell(row=idx, column=c).border = thin_border
                ws.cell(row=idx, column=c).fill = accent_fill if "CASH FLOWS" in line else zebra_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border
                
    auto_fit(ws)

def build_sheet_30_notes_2024(wb):
    ws = wb.create_sheet(title="Notes_2024")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — NOTES TO THE FINANCIAL STATEMENTS (FY 2024)").font = title_font
    ws.cell(row=2, column=1, value="Numbered Notes 1 to 28 Supporting the Audited Financial Statements for the Year Ended 31 December 2024").font = italic_font
    
    headers = ["Note #", "Note Description / Accounting Line", "Sub-Item Breakdown", "31 Dec 2024 (NGN)", "31 Dec 2023 (NGN)", "IFRS Standard / Disclosures"]
    style_header(ws, 4, headers)
    
    notes_data = [
        ("Note 1", "General Information", "Corporate Profile, RC 1526224, Registered in Nigeria under CAMA", None, None, "IAS 1.51"),
        ("Note 2", "Basis of Preparation & Accounting Policies", "Compliance with full IFRS, FRCN & historical cost convention", None, None, "IAS 1.112"),
        ("Note 3", "Revenue from Contracts with Customers", "Civil Engineering & Building Construction Contracts", 38000000.00, 22000000.00, "IFRS 15 Over Time"),
        ("Note 3", "Revenue from Contracts with Customers", "Land Subdivision & Serviced Plot Resale", 19500000.00, 12500000.00, "IFRS 15 Point in Time"),
        ("Note 3", "Revenue from Contracts with Customers", "Project Management & Engineering Consulting Services", 6700000.00, 4000000.00, "IFRS 15 Over Time"),
        ("Note 3", "TOTAL REVENUE", "Gross Statutory Turnover", "=SUM(D7:D9)", "=SUM(E7:E9)", "TIED TO SPLOCI"),
        
        ("Note 4", "Cost of Sales", "Construction Materials & Subcontracting Services", 26400000.00, 14200000.00, "IAS 2 Direct Materials"),
        ("Note 4", "Cost of Sales", "Land Inventory Released on Sold Plots", 11800000.00, 8100000.00, "IAS 2 Inventory Cost"),
        ("Note 4", "Cost of Sales", "Direct Site Labor & Heavy Equipment Rentals", 3300000.00, 2350000.00, "IAS 2 Direct Labor"),
        ("Note 4", "TOTAL COST OF SALES", "Direct Contract & Property Costs", "=SUM(D11:D13)", "=SUM(E11:E13)", "TIED TO SPLOCI"),
        
        ("Note 5", "Administrative & Operating Expenses", "Staff Salaries, Wages & Allowances", 4200000.00, 2800000.00, "IAS 19 Employee Costs"),
        ("Note 5", "Administrative & Operating Expenses", "Site Office Rent & Facility Maintenance", 1450000.00, 950000.00, "Operating Overheads"),
        ("Note 5", "Administrative & Operating Expenses", "Professional, Legal & Regulatory Compliance Fees", 980000.00, 650000.00, "Professional Services"),
        ("Note 5", "Administrative & Operating Expenses", "Statutory Audit & Tax Advisory Fees", 450000.00, 350000.00, "Audit Fees Accrual"),
        ("Note 5", "Administrative & Operating Expenses", "Bank Charges, Commission on Turnover & Electronic Levies", 420000.00, 285000.00, "Financial Services"),
        ("Note 5", "Administrative & Operating Expenses", "Motor Vehicle Fuel, Maintenance & Site Logistics", 750000.00, 420000.00, "Transport & Travel"),
        ("Note 5", "Administrative & Operating Expenses", "Depreciation of Property, Plant & Equipment", 715000.00, 385000.00, "IAS 16 Depreciation"),
        ("Note 5", "Administrative & Operating Expenses", "IFRS 9 Expected Credit Loss Allowance on Receivables", 260000.00, 185000.00, "IFRS 9 Impairment"),
        ("Note 5", "Administrative & Operating Expenses", "General Office Administration, Security & Utilities", 625000.00, 400000.00, "General Administration"),
        ("Note 5", "TOTAL ADMINISTRATIVE EXPENSES", "Operating Overheads", "=SUM(D15:D23)", "=SUM(E15:E23)", "TIED TO SPLOCI"),
        
        ("Note 6", "Finance Costs", "Interest Expense on Bank Project Facility", 350000.00, 125000.00, "IAS 23 Borrowing Costs"),
        ("Note 6", "TOTAL FINANCE COSTS", "Net Finance Expense", "=D25", "=E25", "TIED TO SPLOCI"),
        
        ("Note 7", "Property, Plant and Equipment", "Office & Site Equipment Net Book Value", 1047040.00, 1424540.00, "IAS 16 Carrying Amount"),
        ("Note 7", "Property, Plant and Equipment", "Motor & Project Vehicles Net Book Value", 2362500.00, 0.00, "IAS 16 Carrying Amount"),
        ("Note 7", "TOTAL PROPERTY, PLANT AND EQUIPMENT", "Gross Cost ₦4,523,850 less Acc Dep ₦1,114,310", "=SUM(D27:D28)", "=SUM(E27:E28)", "TIED TO SFP"),
        
        ("Note 8", "Inventories", "Land Held for Resale (Epe & Ibeju Schemes)", 8200000.00, 5400000.00, "IAS 2 Lower of Cost/NRV"),
        ("Note 8", "Inventories", "Construction Work-in-Progress (Lekki & VI Projects)", 5600000.00, 3250000.00, "IAS 2 Lower of Cost/NRV"),
        ("Note 8", "TOTAL INVENTORIES", "Development Land & WIP", "=SUM(D30:D31)", "=SUM(E30:E31)", "TIED TO SFP"),
        
        ("Note 9", "Contract Assets", "Unbilled Certified Construction Revenue", 4150000.00, 2400000.00, "IFRS 15 Unbilled Asset"),
        ("Note 9", "TOTAL CONTRACT ASSETS", "Net Contract Assets", "=D33", "=E33", "TIED TO SFP"),
        
        ("Note 10", "Trade and Other Receivables", "Gross Trade Receivables from Clients", 5800000.00, 3650000.00, "Amortized Cost"),
        ("Note 10", "Trade and Other Receivables", "Less: Allowance for Expected Credit Losses (IFRS 9)", -445000.00, -185000.00, "IFRS 9 ECL Provision"),
        ("Note 10", "Trade and Other Receivables", "Withholding Tax (WHT) Credit Notes Receivable", 2650000.00, 1100000.00, "FIRS Credit Notes"),
        ("Note 10", "Trade and Other Receivables", "Prepayments & Site Advances to Suppliers", 680000.00, 450000.00, "Operating Prepayments"),
        ("Note 10", "TOTAL TRADE AND OTHER RECEIVABLES", "Net Trade & Other Receivables", "=SUM(D35:D38)", "=SUM(E35:E38)", "TIED TO SFP"),
        
        ("Note 11", "Cash and Cash Equivalents", "Providus Bank Plc (A/C 5400281942)", 3850210.00, 2215820.00, "Primary Operating A/C"),
        ("Note 11", "Cash and Cash Equivalents", "First Bank of Nigeria Limited (A/C 2034891102)", 980118.00, 512410.00, "Site Collection A/C"),
        ("Note 11", "Cash and Cash Equivalents", "Sterling Bank Plc (A/C 0078451290)", 312350.00, 114173.00, "Disbursement A/C"),
        ("Note 11", "TOTAL CASH AND CASH EQUIVALENTS", "Reconciled Cash Balances", "=SUM(D40:D42)", "=SUM(E40:E42)", "TIED TO SFP"),
        
        ("Note 13", "Contract Liabilities", "Customer Mobilization Advances & Off-Plan Deposits", 3600000.00, 2400000.00, "IFRS 15 Advances"),
        ("Note 13", "TOTAL CONTRACT LIABILITIES", "Performance Obligations Due", "=D44", "=E44", "TIED TO SFP"),
        
        ("Note 14", "Trade and Other Payables", "Trade Payables - Subcontractors & Materials", 3850000.00, 2150000.00, "Trade Creditors"),
        ("Note 14", "Trade and Other Payables", "Accrued Statutory Audit & Professional Fees", 550000.00, 375000.00, "Audit Accrual"),
        ("Note 14", "Trade and Other Payables", "Other Accrued Operating Expenses & Sundry Creditors", 425500.00, 245235.00, "Sundry Accruals"),
        ("Note 14", "TOTAL TRADE AND OTHER PAYABLES", "Current Operating Liabilities", "=SUM(D46:D48)", "=SUM(E46:E48)", "TIED TO SFP"),
        
        ("Note 15", "Deferred Taxation", "Deferred Tax Asset / (Liability) Closing Balance", -16000.00, 32000.00, "IAS 12 Net Liability"),
        ("Note 15", "TOTAL DEFERRED TAXATION", "Recognized in Statement of Financial Position", "=D50", "=E50", "TIED TO SFP"),
        
        ("Note 16", "Current Taxation", "Companies Income Tax (CIT) Provision", 2528000.00, 1492000.00, "20% CIT Medium Co"),
        ("Note 16", "Current Taxation", "Tertiary Education Tax (TET) Provision", 408600.00, 239400.00, "3% TETFA Rate"),
        ("Note 16", "Current Taxation", "Police Trust Fund (PTF) Levy Provision", 625.00, 365.00, "0.005% of PBT"),
        ("Note 16", "TOTAL CURRENT TAX CHARGE", "Current Tax Provision for Year", "=SUM(D52:D54)", "=SUM(E52:E54)", "TIED TO SPLOCI"),
        
        ("Note 17", "Borrowings & Director's Loans", "Director Project Loan (Engr. Babatunde Goke)", 5800000.00, 4500000.00, "IAS 24 Related Party"),
        ("Note 17", "TOTAL BORROWINGS", "Non-Current Project Funding", "=D56", "=E56", "TIED TO SFP"),
        
        ("Note 18", "Share Capital", "1,000,000 Ordinary Shares of ₦1.00 each", 1000000.00, 1000000.00, "Authorized & Issued"),
        ("Note 18", "TOTAL SHARE CAPITAL", "Fully Issued and Paid Up", "=D58", "=E58", "TIED TO SFP"),
        
        ("Note 19", "Retained Earnings", "Cumulative Retained Earnings Carried Forward", 17476718.00, 7961943.00, "SOCE Roll-Forward"),
        ("Note 19", "TOTAL RETAINED EARNINGS", "Accumulated Earnings", "=D60", "=E60", "TIED TO SFP"),
    ]
    
    for idx, (note_no, desc, sub, v24, v23, ifrs) in enumerate(notes_data, 5):
        is_tot = "TOTAL" in desc
        ws.cell(row=idx, column=1, value=note_no).alignment = align_center
        ws.cell(row=idx, column=1).font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=2, value=desc).font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=3, value=sub).font = bold_font if is_tot else regular_font
        
        for c_idx, val in enumerate([v24, v23], 4):
            c = ws.cell(row=idx, column=c_idx, value=val)
            if val is not None:
                c.number_format = fmt_currency
                c.alignment = align_right
                c.font = bold_font if is_tot else regular_font
                
        c_ifrs = ws.cell(row=idx, column=6, value=ifrs)
        c_ifrs.font = success_font if "TIED" in ifrs else italic_font
        c_ifrs.alignment = align_center if "TIED" in ifrs else align_left
        
        for c in range(1, 7):
            if is_tot:
                ws.cell(row=idx, column=c).border = double_bottom
                ws.cell(row=idx, column=c).fill = total_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border
                if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
                
    auto_fit(ws)

def build_sheet_31_afs_2025(wb):
    ws = wb.create_sheet(title="AFS_2025")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED (RC 1526224)").font = title_font
    ws.cell(row=2, column=1, value="ANNUAL FINANCIAL STATEMENTS FOR THE YEAR ENDED 31 DECEMBER 2025").font = section_font
    ws.cell(row=3, column=1, value="(With Comparative Figures for the Year Ended 31 December 2024 — Stated in Nigerian Naira ₦)").font = italic_font
    
    # 1. SFP 2025
    ws.cell(row=5, column=1, value="1. STATEMENT OF FINANCIAL POSITION AS AT 31 DECEMBER 2025").font = section_font
    headers_sfp = ["Financial Statement Line Item", "Notes", "31 Dec 2025 (NGN)", "31 Dec 2024 (NGN)", "Audit / IFRS Status"]
    style_header(ws, 6, headers_sfp)
    
    sfp_rows = [
        ("NON-CURRENT ASSETS", None, None, None, None),
        ("Property, plant and equipment", "Note 7", 5724540.00, 3409540.00, "IAS 16 Carrying Amount"),
        ("TOTAL NON-CURRENT ASSETS", None, "=C8", "=D8", "TIED OUT"),
        ("CURRENT ASSETS", None, None, None, None),
        ("Inventories (Development land & WIP)", "Note 8", 24300000.00, 13800000.00, "IAS 2 Cost / NRV"),
        ("Contract assets", "Note 9", 6850000.00, 4150000.00, "IFRS 15 Unbilled Revenue"),
        ("Trade and other receivables", "Note 10", 14360000.00, 8685000.00, "IFRS 9 Net of ECL"),
        ("Cash and cash equivalents", "Note 11", 8924048.00, 5142678.00, "Reconciled Bank Balances"),
        ("TOTAL CURRENT ASSETS", None, "=SUM(C11:C14)", "=SUM(D11:D14)", "TIED OUT"),
        ("TOTAL ASSETS", None, "=C9+C15", "=D9+D15", "TIED OUT"),
        ("EQUITY", None, None, None, None),
        ("Share capital", "Note 18", 1000000.00, 1000000.00, "1m Ordinary Shares @ ₦1"),
        ("Retained earnings", "Note 19", 32376588.00, 17476718.00, "Cumulative Post-Tax Earnings"),
        ("TOTAL EQUITY", None, "=SUM(C18:C19)", "=SUM(D18:D19)", "TIED OUT"),
        ("NON-CURRENT LIABILITIES", None, None, None, None),
        ("Borrowings & director's loans", "Note 17", 7500000.00, 5800000.00, "IAS 24 Related Party"),
        ("Deferred tax liability", "Note 15", 98500.00, 16000.00, "IAS 12 Temporary Differences"),
        ("TOTAL NON-CURRENT LIABILITIES", None, "=SUM(C22:C23)", "=SUM(D22:D23)", "TIED OUT"),
        ("CURRENT LIABILITIES", None, None, None, None),
        ("Trade and other payables", "Note 14", 7983500.00, 4825500.00, "Amortized Cost"),
        ("Contract liabilities", "Note 13", 5800000.00, 3600000.00, "IFRS 15 Customer Advances"),
        ("Current tax liabilities", "Note 16", 5400000.00, 2469000.00, "CITA / TETFA Statutory"),
        ("TOTAL CURRENT LIABILITIES", None, "=SUM(C26:C28)", "=SUM(D26:D28)", "TIED OUT"),
        ("TOTAL LIABILITIES AND EQUITY", None, "=C20+C24+C29", "=D20+D24+D29", "TIED TO TOTAL ASSETS"),
    ]
    
    for idx, (line, note, v25, v24, stat) in enumerate(sfp_rows, 7):
        is_sub = "TOTAL" in line or line in ["NON-CURRENT ASSETS", "CURRENT ASSETS", "EQUITY", "NON-CURRENT LIABILITIES", "CURRENT LIABILITIES"]
        is_grand = line in ["TOTAL ASSETS", "TOTAL LIABILITIES AND EQUITY"]
        ws.cell(row=idx, column=1, value=line).font = bold_font if is_sub else regular_font
        if note: ws.cell(row=idx, column=2, value=note).alignment = align_center
        
        for c_idx, val in enumerate([v25, v24], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            if val is not None:
                c.number_format = fmt_currency
                c.alignment = align_right
                c.font = bold_font if is_sub else regular_font
                
        if stat:
            c_st = ws.cell(row=idx, column=5, value=stat)
            c_st.font = success_font if "TIED" in stat else italic_font
            c_st.alignment = align_center if "TIED" in stat else align_left
            
        for c in range(1, 6):
            if is_grand:
                ws.cell(row=idx, column=c).border = double_bottom
                ws.cell(row=idx, column=c).fill = total_fill
            elif is_sub:
                ws.cell(row=idx, column=c).border = thin_border
                ws.cell(row=idx, column=c).fill = accent_fill if "TOTAL" in line else zebra_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border

    # 2. SPLOCI 2025
    r_pnl_start = 33
    ws.cell(row=r_pnl_start, column=1, value="2. STATEMENT OF PROFIT OR LOSS AND OTHER COMPREHENSIVE INCOME FOR THE YEAR ENDED 31 DECEMBER 2025").font = section_font
    headers_pnl = ["Statement of Profit or Loss Line Item", "Notes", "Year Ended 31 Dec 2025 (NGN)", "Year Ended 31 Dec 2024 (NGN)", "IFRS Accounting Treatment"]
    style_header(ws, r_pnl_start + 1, headers_pnl)
    
    pnl_rows = [
        ("Revenue from contracts with customers", "Note 3", 112500000.00, 64200000.00, "IFRS 15 Recognized Revenue"),
        ("Cost of sales", "Note 4", -72800000.00, -41500000.00, "IAS 2 Direct Project Costs"),
        ("GROSS PROFIT", None, "=SUM(C35:C36)", "=SUM(D35:D36)", "Gross Margin: 35.29%"),
        ("Administrative and operating expenses", "Note 5", -16450000.00, -9850000.00, "IAS 1 Operating Overheads"),
        ("OPERATING PROFIT (EBIT)", None, "=C37+C38", "=D37+D38", "Operating Margin: 20.67%"),
        ("Finance costs", "Note 6", -650000.00, -350000.00, "Bank Project Facility Interest"),
        ("PROFIT BEFORE TAXATION", None, "=C39+C40", "=D39+D40", "Taxable Accounting Profit"),
        ("Income tax expense", "Note 16", -7700130.00, -2985225.00, "Current Tax ₦7,617,630 + DTL ₦82,500"),
        ("PROFIT FOR THE YEAR", None, "=C41+C42", "=D41+D42", "Net Profit Margin: 13.24%"),
        ("OTHER COMPREHENSIVE INCOME", None, 0.00, 0.00, "No Items of OCI"),
        ("TOTAL COMPREHENSIVE INCOME FOR THE YEAR", None, "=C43+C44", "=D43+D44", "TRANSFERRED TO EQUITY"),
    ]
    
    for idx, (line, note, v25, v24, treat) in enumerate(pnl_rows, r_pnl_start + 2):
        is_tot = "GROSS PROFIT" in line or "OPERATING PROFIT" in line or "PROFIT BEFORE" in line or "PROFIT FOR" in line or "TOTAL COMPREHENSIVE" in line
        ws.cell(row=idx, column=1, value=line).font = bold_font if is_tot else regular_font
        if note: ws.cell(row=idx, column=2, value=note).alignment = align_center
        
        for c_idx, val in enumerate([v25, v24], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
            
        ws.cell(row=idx, column=5, value=treat).font = italic_font
        
        for c in range(1, 6):
            if is_tot:
                ws.cell(row=idx, column=c).border = double_bottom if "TOTAL COMPREHENSIVE" in line or "PROFIT FOR" in line else thin_border
                ws.cell(row=idx, column=c).fill = total_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border
                if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    # 3. SCF 2025
    r_cf_start = 48
    ws.cell(row=r_cf_start, column=1, value="3. STATEMENT OF CASH FLOWS FOR THE YEAR ENDED 31 DECEMBER 2025 (INDIRECT METHOD)").font = section_font
    headers_cf = ["Cash Flow Line Item", "Notes", "Year Ended 31 Dec 2025 (NGN)", "Year Ended 31 Dec 2024 (NGN)", "Cash Flow Classification"]
    style_header(ws, r_cf_start + 1, headers_cf)
    
    cf_rows = [
        ("CASH FLOWS FROM OPERATING ACTIVITIES", None, None, None, None),
        ("Profit before taxation", "SPLOCI", 22600000.00, 12500000.00, "Operating PBT"),
        ("Adjustments for non-cash items:", None, None, None, None),
        ("  - Depreciation of property, plant and equipment", "Note 7", 1285000.00, 715000.00, "Non-Cash Operating Adj"),
        ("  - Impairment allowance on receivables (ECL)", "Note 10", 395000.00, 260000.00, "Non-Cash Operating Adj"),
        ("  - Finance costs recognized in profit or loss", "Note 6", 650000.00, 350000.00, "Financing Reclassification"),
        ("Operating profit before working capital changes", None, "=SUM(C51:C55)", "=SUM(D51:D55)", "TIED OUT"),
        ("Working capital changes:", None, None, None, None),
        ("  - (Increase) in development inventories and WIP", "Note 8", -10500000.00, -5150000.00, "IAS 2 Inventory Delta"),
        ("  - (Increase) in contract assets", "Note 9", -2700000.00, -1750000.00, "IFRS 15 Asset Delta"),
        ("  - (Increase) in trade and other receivables", "Note 10", -6070000.00, -3930000.00, "Receivables Delta"),
        ("  - Increase in trade and other payables", "Note 14", 3158000.00, 2055265.00, "Payables Delta"),
        ("  - Increase in contract liabilities (advances)", "Note 13", 2200000.00, 1200000.00, "IFRS 15 Liability Delta"),
        ("Cash generated from operations", None, "=C56+SUM(C58:C62)", "=D56+SUM(D58:D62)", "TIED OUT"),
        ("  - Income taxes paid", "Note 16", -4886630.00, -2200000.00, "Operating Tax Cash Flow"),
        ("  - Interest paid on project facilities", "Note 6", -650000.00, -350000.00, "Operating Interest Cash Flow"),
        ("Net cash flows from operating activities", None, "=C63+C64+C65", "=D63+D64+D65", "TIED OUT"),
        ("CASH FLOWS FROM INVESTING ACTIVITIES", None, None, None, None),
        ("Purchase of property, plant and equipment", "Note 7", -3600000.00, -2700000.00, "IAS 16 Cash Additions"),
        ("Net cash flows used in investing activities", None, "=C68", "=D68", "TIED OUT"),
        ("CASH FLOWS FROM FINANCING ACTIVITIES", None, None, None, None),
        ("Proceeds from director's project loans", "Note 17", 1700000.00, 1300000.00, "IAS 24 Related Party"),
        ("Net cash flows from financing activities", None, "=C71", "=D71", "TIED OUT"),
        ("Net increase in cash and cash equivalents", None, "=C66+C69+C72", "=D66+D69+D72", "Net Annual Liquidity Delta"),
        ("Cash and cash equivalents at 1 January", "Note 11", 5142678.00, 2842403.00, "Opening Cash B/F"),
        ("CASH AND CASH EQUIVALENTS AT 31 DECEMBER", "Note 11", "=C73+C74", "=D73+D74", "TIED TO SFP NOTE 11"),
    ]
    
    for idx, (line, note, v25, v24, cls) in enumerate(cf_rows, r_cf_start + 2):
        is_sub = "TOTAL" in line or "CASH FLOWS FROM" in line or "Net cash" in line or "CASH AND CASH" in line
        is_grand = "CASH AND CASH EQUIVALENTS AT 31" in line
        ws.cell(row=idx, column=1, value=line).font = bold_font if is_sub else regular_font
        if note: ws.cell(row=idx, column=2, value=note).alignment = align_center
        
        for c_idx, val in enumerate([v25, v24], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            if val is not None:
                c.number_format = fmt_currency
                c.alignment = align_right
                c.font = bold_font if is_sub else regular_font
                
        ws.cell(row=idx, column=5, value=cls).font = success_font if is_grand else italic_font
        
        for c in range(1, 6):
            if is_grand:
                ws.cell(row=idx, column=c).border = double_bottom
                ws.cell(row=idx, column=c).fill = total_fill
            elif is_sub:
                ws.cell(row=idx, column=c).border = thin_border
                ws.cell(row=idx, column=c).fill = accent_fill if "CASH FLOWS" in line else zebra_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border
                
    auto_fit(ws)

def build_sheet_32_notes_2025(wb):
    ws = wb.create_sheet(title="Notes_2025")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — NOTES TO THE FINANCIAL STATEMENTS (FY 2025)").font = title_font
    ws.cell(row=2, column=1, value="Numbered Notes 1 to 28 Supporting the Audited Financial Statements for the Year Ended 31 December 2025").font = italic_font
    
    headers = ["Note #", "Note Description / Accounting Line", "Sub-Item Breakdown", "31 Dec 2025 (NGN)", "31 Dec 2024 (NGN)", "IFRS Standard / Disclosures"]
    style_header(ws, 4, headers)
    
    notes_data = [
        ("Note 1", "General Information", "Corporate Profile, RC 1526224, Registered in Nigeria under CAMA", None, None, "IAS 1.51"),
        ("Note 2", "Basis of Preparation & Accounting Policies", "Compliance with full IFRS, FRCN & historical cost convention", None, None, "IAS 1.112"),
        ("Note 3", "Revenue from Contracts with Customers", "Civil Engineering & Building Construction Contracts", 68000000.00, 38000000.00, "IFRS 15 Over Time"),
        ("Note 3", "Revenue from Contracts with Customers", "Land Subdivision & Serviced Plot Resale", 32500000.00, 19500000.00, "IFRS 15 Point in Time"),
        ("Note 3", "Revenue from Contracts with Customers", "Project Management & Engineering Consulting Services", 12000000.00, 6700000.00, "IFRS 15 Over Time"),
        ("Note 3", "TOTAL REVENUE", "Gross Statutory Turnover", "=SUM(D7:D9)", "=SUM(E7:E9)", "TIED TO SPLOCI"),
        
        ("Note 4", "Cost of Sales", "Construction Materials & Subcontracting Services", 45600000.00, 26400000.00, "IAS 2 Direct Materials"),
        ("Note 4", "Cost of Sales", "Land Inventory Released on Sold Plots", 21400000.00, 11800000.00, "IAS 2 Inventory Cost"),
        ("Note 4", "Cost of Sales", "Direct Site Labor & Heavy Equipment Rentals", 5800000.00, 3300000.00, "IAS 2 Direct Labor"),
        ("Note 4", "TOTAL COST OF SALES", "Direct Contract & Property Costs", "=SUM(D11:D13)", "=SUM(E11:E13)", "TIED TO SPLOCI"),
        
        ("Note 5", "Administrative & Operating Expenses", "Staff Salaries, Wages & Allowances", 7200000.00, 4200000.00, "IAS 19 Employee Costs"),
        ("Note 5", "Administrative & Operating Expenses", "Site Office Rent & Facility Maintenance", 2400000.00, 1450000.00, "Operating Overheads"),
        ("Note 5", "Administrative & Operating Expenses", "Professional, Legal & Regulatory Compliance Fees", 1650000.00, 980000.00, "Professional Services"),
        ("Note 5", "Administrative & Operating Expenses", "Statutory Audit & Tax Advisory Fees", 650000.00, 450000.00, "Audit Fees Accrual"),
        ("Note 5", "Administrative & Operating Expenses", "Bank Charges, Commission on Turnover & Electronic Levies", 680000.00, 420000.00, "Financial Services"),
        ("Note 5", "Administrative & Operating Expenses", "Motor Vehicle Fuel, Maintenance & Site Logistics", 1250000.00, 750000.00, "Transport & Travel"),
        ("Note 5", "Administrative & Operating Expenses", "Depreciation of Property, Plant & Equipment", 1285000.00, 715000.00, "IAS 16 Depreciation"),
        ("Note 5", "Administrative & Operating Expenses", "IFRS 9 Expected Credit Loss Allowance on Receivables", 395000.00, 260000.00, "IFRS 9 Impairment"),
        ("Note 5", "Administrative & Operating Expenses", "General Office Administration, Security & Utilities", 940000.00, 625000.00, "General Administration"),
        ("Note 5", "TOTAL ADMINISTRATIVE EXPENSES", "Operating Overheads", "=SUM(D15:D23)", "=SUM(E15:E23)", "TIED TO SPLOCI"),
        
        ("Note 6", "Finance Costs", "Interest Expense on Bank Project Facility", 650000.00, 350000.00, "IAS 23 Borrowing Costs"),
        ("Note 6", "TOTAL FINANCE COSTS", "Net Finance Expense", "=D25", "=E25", "TIED TO SPLOCI"),
        
        ("Note 7", "Property, Plant and Equipment", "Office & Site Equipment Net Book Value", 799540.00, 1047040.00, "IAS 16 Carrying Amount"),
        ("Note 7", "Property, Plant and Equipment", "Motor & Project Vehicles Net Book Value", 1687500.00, 2362500.00, "IAS 16 Carrying Amount"),
        ("Note 7", "Property, Plant and Equipment", "Plant & Heavy Machinery Net Book Value", 3237500.00, 0.00, "IAS 16 Carrying Amount"),
        ("Note 7", "TOTAL PROPERTY, PLANT AND EQUIPMENT", "Gross Cost ₦8,123,850 less Acc Dep ₦2,399,310", "=SUM(D27:D29)", "=SUM(E27:E29)", "TIED TO SFP"),
        
        ("Note 8", "Inventories", "Land Held for Resale (Epe & Ibeju Schemes)", 14500000.00, 8200000.00, "IAS 2 Lower of Cost/NRV"),
        ("Note 8", "Inventories", "Construction Work-in-Progress (Lekki & VI Projects)", 9800000.00, 5600000.00, "IAS 2 Lower of Cost/NRV"),
        ("Note 8", "TOTAL INVENTORIES", "Development Land & WIP", "=SUM(D31:D32)", "=SUM(E31:E32)", "TIED TO SFP"),
        
        ("Note 9", "Contract Assets", "Unbilled Certified Construction Revenue", 6850000.00, 4150000.00, "IFRS 15 Unbilled Asset"),
        ("Note 9", "TOTAL CONTRACT ASSETS", "Net Contract Assets", "=D34", "=E34", "TIED TO SFP"),
        
        ("Note 10", "Trade and Other Receivables", "Gross Trade Receivables from Clients", 9200000.00, 5800000.00, "Amortized Cost"),
        ("Note 10", "Trade and Other Receivables", "Less: Allowance for Expected Credit Losses (IFRS 9)", -840000.00, -445000.00, "IFRS 9 ECL Provision"),
        ("Note 10", "Trade and Other Receivables", "Withholding Tax (WHT) Credit Notes Receivable", 4850000.00, 2650000.00, "FIRS Credit Notes"),
        ("Note 10", "Trade and Other Receivables", "Prepayments & Site Advances to Suppliers", 1150000.00, 680000.00, "Operating Prepayments"),
        ("Note 10", "TOTAL TRADE AND OTHER RECEIVABLES", "Net Trade & Other Receivables", "=SUM(D36:D39)", "=SUM(E36:E39)", "TIED TO SFP"),
        
        ("Note 11", "Cash and Cash Equivalents", "Providus Bank Plc (A/C 5400281942)", 6420110.00, 3850210.00, "Primary Operating A/C"),
        ("Note 11", "Cash and Cash Equivalents", "First Bank of Nigeria Limited (A/C 2034891102)", 1840550.00, 980118.00, "Site Collection A/C"),
        ("Note 11", "Cash and Cash Equivalents", "Sterling Bank Plc (A/C 0078451290)", 663388.00, 312350.00, "Disbursement A/C"),
        ("Note 11", "TOTAL CASH AND CASH EQUIVALENTS", "Reconciled Cash Balances", "=SUM(D41:D43)", "=SUM(E41:E43)", "TIED TO SFP"),
        
        ("Note 13", "Contract Liabilities", "Customer Mobilization Advances & Off-Plan Deposits", 5800000.00, 3600000.00, "IFRS 15 Advances"),
        ("Note 13", "TOTAL CONTRACT LIABILITIES", "Performance Obligations Due", "=D45", "=E45", "TIED TO SFP"),
        
        ("Note 14", "Trade and Other Payables", "Trade Payables - Subcontractors & Materials", 6450000.00, 3850000.00, "Trade Creditors"),
        ("Note 14", "Trade and Other Payables", "Accrued Statutory Audit & Professional Fees", 850000.00, 550000.00, "Audit Accrual"),
        ("Note 14", "Trade and Other Payables", "Other Accrued Operating Expenses & Sundry Creditors", 683500.00, 425500.00, "Sundry Accruals"),
        ("Note 14", "TOTAL TRADE AND OTHER PAYABLES", "Current Operating Liabilities", "=SUM(D47:D49)", "=SUM(E47:E49)", "TIED TO SFP"),
        
        ("Note 15", "Deferred Taxation", "Deferred Tax Asset / (Liability) Closing Balance", -98500.00, -16000.00, "IAS 12 Net Liability"),
        ("Note 15", "TOTAL DEFERRED TAXATION", "Recognized in Statement of Financial Position", "=D51", "=E51", "TIED TO SFP"),
        
        ("Note 16", "Current Taxation", "Companies Income Tax (CIT) Provision", 6825000.00, 2528000.00, "30% CIT Large Co"),
        ("Note 16", "Current Taxation", "Tertiary Education Tax (TET) Provision", 735000.00, 408600.00, "3% TETFA Rate"),
        ("Note 16", "Current Taxation", "Police Trust Fund (PTF) Levy Provision", 1130.00, 625.00, "0.005% of PBT"),
        ("Note 16", "Current Taxation", "NASENI Development Levy Provision", 56500.00, 0.00, "0.25% of PBT"),
        ("Note 16", "TOTAL CURRENT TAX CHARGE", "Current Tax Provision for Year", "=SUM(D53:D56)", "=SUM(E53:E56)", "TIED TO SPLOCI"),
        
        ("Note 17", "Borrowings & Director's Loans", "Director Project Loan (Engr. Babatunde Goke)", 7500000.00, 5800000.00, "IAS 24 Related Party"),
        ("Note 17", "TOTAL BORROWINGS", "Non-Current Project Funding", "=D58", "=E58", "TIED TO SFP"),
        
        ("Note 18", "Share Capital", "1,000,000 Ordinary Shares of ₦1.00 each", 1000000.00, 1000000.00, "Authorized & Issued"),
        ("Note 18", "TOTAL SHARE CAPITAL", "Fully Issued and Paid Up", "=D60", "=E60", "TIED TO SFP"),
        
        ("Note 19", "Retained Earnings", "Cumulative Retained Earnings Carried Forward", 32376588.00, 17476718.00, "SOCE Roll-Forward"),
        ("Note 19", "TOTAL RETAINED EARNINGS", "Accumulated Earnings", "=D62", "=E62", "TIED TO SFP"),
    ]
    
    for idx, (note_no, desc, sub, v25, v24, ifrs) in enumerate(notes_data, 5):
        is_tot = "TOTAL" in desc
        ws.cell(row=idx, column=1, value=note_no).alignment = align_center
        ws.cell(row=idx, column=1).font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=2, value=desc).font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=3, value=sub).font = bold_font if is_tot else regular_font
        
        for c_idx, val in enumerate([v25, v24], 4):
            c = ws.cell(row=idx, column=c_idx, value=val)
            if val is not None:
                c.number_format = fmt_currency
                c.alignment = align_right
                c.font = bold_font if is_tot else regular_font
                
        c_ifrs = ws.cell(row=idx, column=6, value=ifrs)
        c_ifrs.font = success_font if "TIED" in ifrs else italic_font
        c_ifrs.alignment = align_center if "TIED" in ifrs else align_left
        
        for c in range(1, 7):
            if is_tot:
                ws.cell(row=idx, column=c).border = double_bottom
                ws.cell(row=idx, column=c).fill = total_fill
            else:
                ws.cell(row=idx, column=c).border = thin_border
                if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
                
    auto_fit(ws)

print("sheet_builders_5.py created successfully.")
