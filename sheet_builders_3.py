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

def build_sheet_15_project_register(wb):
    ws = wb.create_sheet(title="Project_Register")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — COMPREHENSIVE PROJECT REGISTER").font = title_font
    ws.cell(row=2, column=1, value="Civil Engineering, Land Resale, Joint Venture Real Estate & Consultancy Portfolio (2023 - 2025)").font = italic_font
    
    headers = ["Project Code", "Project Name & Location", "Project Type", "Landowner / Client", "Legal Rights & Contract Basis", "Contract Value (NGN)", "Commencement", "Target Completion", "Progress / Stage (%)", "Incurred Costs (NGN)", "Forecast Cost to Complete (NGN)", "Billed Revenue (NGN)", "Cash Collections (NGN)", "Retentions / Contract Assets (NGN)", "NRV Test / Valuation Status", "Risk Assessment"]
    style_header(ws, 4, headers)
    
    projects = [
        ("PROJ-101", "Lekki Residential Estate Infrastructure & Civil Works", "Civil Engineering & Building", "Apex Properties Ltd", "Subcontractor Agreement / Enforceable Right to Payment", 65000000.00, "2023-01-10", "2024-12-31", 0.95, 40600000.00, 2140000.00, 60000000.00, 56350000.00, 3650000.00, "Cost < Recoverable Amount (Pass)", "Low / Active Execution"),
        ("PROJ-102", "Ibeju-Lekki Scheme 1 Land Subdivision & Resale", "Land Purchase & Resale", "Ibeju Family Land Committee", "Deed of Assignment & Registered Survey (20 Plots)", 35000000.00, "2023-02-01", "2024-06-30", 1.00, 19900000.00, 0.00, 32000000.00, 32000000.00, 0.00, "Land Sold > Cost (Pass)", "Completed / Fully Sold"),
        ("PROJ-103", "Ikeja Commercial Office Renovation & Civil Works", "Commercial Civil Works", "Horizon Holdings Ltd", "Civil Works Contract & Performance Bond", 1850000.00, "2023-03-01", "2023-12-15", 1.00, 11500000.00, 0.00, 18500000.00, 18500000.00, 0.00, "Completed & Certified (Pass)", "Completed / Fully Settled"),
        ("PROJ-104", "Victoria Island Luxury Maisonettes Development", "Joint Venture Real Estate", "VI Prime Properties & Landowners", "Unincorporated Joint Venture Agreement (40% Profit Share)", 120000000.00, "2024-03-01", "2026-06-30", 0.45, 45200000.00, 55000000.00, 48000000.00, 42000000.00, 6000000.00, "Market Appraisal > Cost (Pass)", "Medium / JV Progress"),
        ("PROJ-105", "Epe Mixed-Use Residential Scheme (Phase 1)", "Land Subdivision & Infrastructure", "Epe Land Registry / Community", "Perimeter Survey & Governor's Consent In Progress", 45000000.00, "2023-11-01", "2025-12-31", 0.60, 23500000.00, 15000000.00, 28000000.00, 24500000.00, 3500000.00, "Valuation > Cost (Pass)", "Low / Strong Demand"),
        ("PROJ-106", "Project Management & Structural Consultancy", "Engineering Consultancy", "Various Corporate Clients", "Annual Engineering Retainers & SLA Contracts", 25000000.00, "2023-01-01", "2025-12-31", 1.00, 1200000.00, 0.00, 22700000.00, 22700000.00, 0.00, "Consultancy Services (Pass)", "Low / Recurring Cashflow"),
    ]
    
    for idx, (pcode, pname, ptype, client, legal, cval, sdt, edt, prg, inc_c, ctc, bill_r, coll_c, ret, nrv, rsk) in enumerate(projects, 5):
        ws.cell(row=idx, column=1, value=pcode).alignment = align_center
        ws.cell(row=idx, column=2, value=pname).font = bold_font
        ws.cell(row=idx, column=3, value=ptype).font = regular_font
        ws.cell(row=idx, column=4, value=client).font = regular_font
        ws.cell(row=idx, column=5, value=legal).font = italic_font
        
        c_cval = ws.cell(row=idx, column=6, value=cval)
        c_cval.number_format = fmt_currency
        c_cval.alignment = align_right
        
        ws.cell(row=idx, column=7, value=sdt).alignment = align_center
        ws.cell(row=idx, column=8, value=edt).alignment = align_center
        
        c_prg = ws.cell(row=idx, column=9, value=prg)
        c_prg.number_format = fmt_percent
        c_prg.alignment = align_right
        
        c_inc = ws.cell(row=idx, column=10, value=inc_c)
        c_ctc = ws.cell(row=idx, column=11, value=ctc)
        c_bill = ws.cell(row=idx, column=12, value=bill_r)
        c_coll = ws.cell(row=idx, column=13, value=coll_c)
        c_ret = ws.cell(row=idx, column=14, value=ret)
        
        for c in [c_inc, c_ctc, c_bill, c_coll, c_ret]:
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = regular_font
            
        ws.cell(row=idx, column=15, value=nrv).font = success_font
        ws.cell(row=idx, column=15).alignment = align_center
        ws.cell(row=idx, column=16, value=rsk).font = regular_font
        ws.cell(row=idx, column=16).alignment = align_center
        
        for c in range(1, 17):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    tot_r = len(projects) + 5
    ws.cell(row=tot_r, column=1, value="TOTAL ACTIVE PORTFOLIO").font = bold_font
    for c_idx, col_let in zip([6, 10, 11, 12, 13, 14], ["F", "J", "K", "L", "M", "N"]):
        c = ws.cell(row=tot_r, column=c_idx, value=f"=SUM({col_let}5:{col_let}{tot_r-1})")
        c.font = bold_font
        c.number_format = fmt_currency
        c.alignment = align_right
    for c in range(1, 17):
        ws.cell(row=tot_r, column=c).border = double_bottom
        ws.cell(row=tot_r, column=c).fill = total_fill
        
    auto_fit(ws)

def build_sheet_16_ifrs15_revenue_wip(wb):
    ws = wb.create_sheet(title="IFRS15_Revenue_WIP")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — IFRS 15 REVENUE & IAS 2 INVENTORY / WIP MODEL").font = title_font
    ws.cell(row=2, column=1, value="5-Step Revenue Recognition Evaluation, Contract Assets/Liabilities & IAS 2 Inventory Roll-Forward").font = italic_font
    
    headers = ["Revenue / Project Stream", "IFRS 15 Timing & Model", "FY 2022 (NGN)", "FY 2023 (NGN)", "FY 2024 (NGN)", "FY 2025 (NGN)", "Accounting Policy & Recognition Criteria"]
    style_header(ws, 4, headers)
    
    rev_streams = [
        ("Civil Engineering & Construction Contracts", "Over Time (Cost-to-Cost Input Method)", 0.00, 22000000.00, 38000000.00, 68000000.00, "Recognized over time as performance creates/enhances asset controlled by client with enforceable payment rights (IFRS 15.35(b))."),
        ("Land Subdivision & Real Estate Sales", "Point in Time (Title Execution / Delivery)", 0.00, 12500000.00, 19500000.00, 32500000.00, "Recognized at point in time when legal deed is executed, physical possession transferred and consideration received/unconditional."),
        ("Project Management & Engineering Consulting", "Over Time (Services Rendered)", 0.00, 4000000.00, 6700000.00, 12000000.00, "Recognized over time as services are simultaneously received and consumed by client (IFRS 15.35(a))."),
        ("TOTAL REVENUE FROM CONTRACTS WITH CUSTOMERS", "Full IFRS 15 Framework", "=SUM(C5:C7)", "=SUM(D5:D7)", "=SUM(E5:E7)", "=SUM(F5:F7)", "Gross statutory turnover recognized in Statement of Profit or Loss."),
    ]
    
    for idx, (stream, model, y22, y23, y24, y25, pol) in enumerate(rev_streams, 5):
        is_tot = "TOTAL" in stream
        ws.cell(row=idx, column=1, value=stream).font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=2, value=model).font = bold_font if is_tot else italic_font
        
        for c_idx, val in enumerate([y22, y23, y24, y25], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
            
        ws.cell(row=idx, column=7, value=pol).font = regular_font
        
        for c in range(1, 8):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    # Contract Assets and Liabilities Table
    ws.cell(row=11, column=1, value="CONTRACT BALANCES ROLL-FORWARD (IFRS 15)").font = section_font
    sub_headers = ["Contract Balance Category", "31 Dec 2022 (NGN)", "31 Dec 2023 (NGN)", "31 Dec 2024 (NGN)", "31 Dec 2025 (NGN)", "Accounting Definition & Nature"]
    style_header(ws, 12, sub_headers)
    
    contract_data = [
        ("Contract Assets (Unbilled Revenue Certified)", 0.00, 2400000.00, 4150000.00, 6850000.00, "Cumulative unbilled revenue for performance completed to date exceeding progress billings issued."),
        ("Contract Liabilities (Customer Mobilization Advances)", 0.00, 2400000.00, 3600000.00, 5800000.00, "Customer advances and deposits received prior to satisfying performance obligations."),
        ("Net Contract Position", "=B13-B14", "=C13-C14", "=D13-D14", "=E13-E14", "Net contract asset / (liability) position under IFRS 15."),
    ]
    
    for idx, (cat, b22, b23, b24, b25, desc) in enumerate(contract_data, 13):
        is_net = "Net Contract" in cat
        ws.cell(row=idx, column=1, value=cat).font = bold_font if is_net else regular_font
        for c_idx, val in enumerate([b22, b23, b24, b25], 2):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_net else regular_font
        ws.cell(row=idx, column=6, value=desc).font = italic_font
        for c in range(1, 7):
            ws.cell(row=idx, column=c).border = double_bottom if is_net else thin_border
            if is_net: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    # IAS 2 Development Inventory & WIP Roll-Forward
    ws.cell(row=18, column=1, value="IAS 2 DEVELOPMENT INVENTORIES & WORK-IN-PROGRESS ROLL-FORWARD").font = section_font
    inv_headers = ["Inventory Component", "Opening Balance 1 Jan (NGN)", "Land & Development Additions (NGN)", "Direct Construction Incurred (NGN)", "Costs Released to Cost of Sales (NGN)", "Closing Balance 31 Dec (NGN)", "NRV Test Status"]
    style_header(ws, 19, inv_headers)
    
    inv_rows = [
        ("FY 2023 — Land Held for Resale (IAS 2)", 0.00, 13500000.00, 0.00, -8100000.00, 5400000.00, "PASSED (NRV > Cost)"),
        ("FY 2023 — Construction WIP (IAS 2)", 0.00, 0.00, 17450000.00, -14200000.00, 3250000.00, "PASSED (NRV > Cost)"),
        ("FY 2023 — TOTAL INVENTORIES", 0.00, 13500000.00, 17450000.00, -22300000.00, 8650000.00, "TIED TO SFP NOTE 8"),
        
        ("FY 2024 — Land Held for Resale (IAS 2)", 5400000.00, 14600000.00, 0.00, -11800000.00, 8200000.00, "PASSED (NRV > Cost)"),
        ("FY 2024 — Construction WIP (IAS 2)", 3250000.00, 0.00, 28750000.00, -26400000.00, 5600000.00, "PASSED (NRV > Cost)"),
        ("FY 2024 — TOTAL INVENTORIES", 8650000.00, 14600000.00, 28750000.00, -38200000.00, 13800000.00, "TIED TO SFP NOTE 8"),
        
        ("FY 2025 — Land Held for Resale (IAS 2)", 8200000.00, 27700000.00, 0.00, -21400000.00, 14500000.00, "PASSED (NRV > Cost)"),
        ("FY 2025 — Construction WIP (IAS 2)", 5600000.00, 0.00, 49800000.00, -45600000.00, 9800000.00, "PASSED (NRV > Cost)"),
        ("FY 2025 — TOTAL INVENTORIES", 13800000.00, 27700000.00, 49800000.00, -67000000.00, 24300000.00, "TIED TO SFP NOTE 8"),
    ]
    
    for idx, (comp, op, add_l, add_c, rel_c, cl, nrv_st) in enumerate(inv_rows, 20):
        is_sub = "TOTAL" in comp
        ws.cell(row=idx, column=1, value=comp).font = bold_font if is_sub else regular_font
        for c_idx, val in enumerate([op, add_l, add_c, rel_c, cl], 2):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_sub else regular_font
        c_st = ws.cell(row=idx, column=7, value=nrv_st)
        c_st.font = success_font if "PASSED" in nrv_st or "TIED" in nrv_st else regular_font
        c_st.alignment = align_center
        for c in range(1, 8):
            ws.cell(row=idx, column=c).border = double_bottom if is_sub else thin_border
            if is_sub: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

def build_sheet_17_ppe_depreciation(wb):
    ws = wb.create_sheet(title="PPE_Depreciation")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — PROPERTY, PLANT AND EQUIPMENT SCHEDULE (IAS 16)").font = title_font
    ws.cell(row=2, column=1, value="Fixed Asset Movement, Additions, Straight-Line Depreciation & Net Book Value (2022 - 2025)").font = italic_font
    
    headers = ["Asset Category", "Depr Rate (%)", "Cost at 31 Dec 2022 (NGN)", "2023 Additions (NGN)", "Cost at 31 Dec 2023 (NGN)", "2024 Additions (NGN)", "Cost at 31 Dec 2024 (NGN)", "2025 Additions (NGN)", "Cost at 31 Dec 2025 (NGN)"]
    style_header(ws, 4, headers)
    
    cost_data = [
        ("Office & Site Equipment", 0.25, 23850.00, 1800000.00, 1823850.00, 0.00, 1823850.00, 0.00, 1823850.00),
        ("Motor & Project Vehicles", 0.25, 0.00, 0.00, 0.00, 2700000.00, 2700000.00, 0.00, 2700000.00),
        ("Plant & Heavy Construction Machinery", 0.20, 0.00, 0.00, 0.00, 0.00, 0.00, 3600000.00, 3600000.00),
        ("TOTAL GROSS PPE COST", None, "=SUM(C5:C7)", "=SUM(D5:D7)", "=SUM(E5:E7)", "=SUM(F5:F7)", "=SUM(G5:G7)", "=SUM(H5:H7)", "=SUM(I5:I7)"),
    ]
    
    for idx, (cat, rate, c22, a23, c23, a24, c24, a25, c25) in enumerate(cost_data, 5):
        is_tot = "TOTAL" in cat
        ws.cell(row=idx, column=1, value=cat).font = bold_font if is_tot else regular_font
        if rate:
            c_r = ws.cell(row=idx, column=2, value=rate)
            c_r.number_format = fmt_percent
            c_r.alignment = align_center
        else:
            ws.cell(row=idx, column=2, value="").alignment = align_center
            
        for c_idx, val in enumerate([c22, a23, c23, a24, c24, a25, c25], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
            
        for c in range(1, 10):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    # Accumulated Depreciation Section
    ws.cell(row=11, column=1, value="ACCUMULATED DEPRECIATION (IAS 16)").font = section_font
    dep_headers = ["Asset Category", "Acc Dep 31 Dec 2022 (NGN)", "2023 Charge (NGN)", "Acc Dep 31 Dec 2023 (NGN)", "2024 Charge (NGN)", "Acc Dep 31 Dec 2024 (NGN)", "2025 Charge (NGN)", "Acc Dep 31 Dec 2025 (NGN)"]
    style_header(ws, 12, dep_headers)
    
    dep_data = [
        ("Office & Site Equipment", 14310.00, 385000.00, 399310.00, 377500.00, 776810.00, 247500.00, 1024310.00),
        ("Motor & Project Vehicles", 0.00, 0.00, 0.00, 337500.00, 337500.00, 675000.00, 1012500.00),
        ("Plant & Heavy Construction Machinery", 0.00, 0.00, 0.00, 0.00, 0.00, 362500.00, 362500.00),
        ("TOTAL ACCUMULATED DEPRECIATION", "=SUM(B13:B15)", "=SUM(C13:C15)", "=SUM(D13:D15)", "=SUM(E13:E15)", "=SUM(F13:F15)", "=SUM(G13:G15)", "=SUM(H13:H15)"),
    ]
    
    for idx, (cat, d22, ch23, d23, ch24, d24, ch25, d25) in enumerate(dep_data, 13):
        is_tot = "TOTAL" in cat
        ws.cell(row=idx, column=1, value=cat).font = bold_font if is_tot else regular_font
        for c_idx, val in enumerate([d22, ch23, d23, ch24, d24, ch25, d25], 2):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
        for c in range(1, 9):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    # Net Book Value Summary Table
    ws.cell(row=19, column=1, value="NET BOOK VALUE (CARRYING AMOUNT)").font = section_font
    nbv_headers = ["Asset Category", "31 Dec 2022 (NGN)", "31 Dec 2023 (NGN)", "31 Dec 2024 (NGN)", "31 Dec 2025 (NGN)", "IAS 16 Compliance Status"]
    style_header(ws, 20, nbv_headers)
    
    nbv_data = [
        ("Office & Site Equipment", 9540.00, 1424540.00, 1047040.00, 799540.00, "Depreciated at 25% Straight-Line"),
        ("Motor & Project Vehicles", 0.00, 0.00, 2362500.00, 1687500.00, "Depreciated at 25% Straight-Line"),
        ("Plant & Heavy Construction Machinery", 0.00, 0.00, 0.00, 3237500.00, "Depreciated at 20% Straight-Line"),
        ("TOTAL PPE NET BOOK VALUE", "=SUM(B21:B23)", "=SUM(C21:C23)", "=SUM(D21:D23)", "=SUM(E21:E23)", "TIED TO SFP NOTE 7"),
    ]
    
    for idx, (cat, n22, n23, n24, n25, stat) in enumerate(nbv_data, 21):
        is_tot = "TOTAL" in cat
        ws.cell(row=idx, column=1, value=cat).font = bold_font if is_tot else regular_font
        for c_idx, val in enumerate([n22, n23, n24, n25], 2):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
        c_st = ws.cell(row=idx, column=6, value=stat)
        c_st.font = success_font if "TIED" in stat else italic_font
        c_st.alignment = align_center
        for c in range(1, 7):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

def build_sheet_18_receivables_ecl_payables(wb):
    ws = wb.create_sheet(title="Receivables_ECL_Payables")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — RECEIVABLES, ECL (IFRS 9) & PAYABLES SUBLEDGER").font = title_font
    ws.cell(row=2, column=1, value="Aging Analysis, IFRS 9 Expected Credit Loss Matrix & Trade Payables Subledger").font = italic_font
    
    headers = ["Debtor / Customer Account", "Current (0-30 days)", "31-60 days", "61-90 days", "Over 90 days", "Gross Carrying Amount (NGN)", "ECL Loss Rate (%)", "ECL Impairment Allowance (NGN)", "Net Carrying Amount (NGN)"]
    style_header(ws, 4, headers)
    
    debtors_2023 = [
        ("Apex Properties Ltd (Lekki Phase 1)", 2500000.00, 0.00, 0.00, 0.00, 2500000.00, 0.02, 50000.00, 2450000.00),
        ("Horizon Holdings Ltd (Ikeja Commercial)", 0.00, 650000.00, 0.00, 0.00, 650000.00, 0.05, 32500.00, 617500.00),
        ("Private Land Buyers (Ibeju Scheme 1)", 0.00, 0.00, 500000.00, 0.00, 500000.00, 0.15, 75000.00, 425000.00),
        ("Historical 2022 Contract Retentions", 0.00, 0.00, 0.00, 0.00, 0.00, 0.25, 27500.00, -27500.00),
        ("TOTAL TRADE RECEIVABLES 2023", "=SUM(B5:B8)", "=SUM(C5:C8)", "=SUM(D5:D8)", "=SUM(E5:E8)", "=SUM(F5:F8)", None, "=SUM(H5:H8)", "=F9-H9"),
    ]
    
    for idx, (deb, c0, c30, c60, c90, gross, lrate, ecl, net) in enumerate(debtors_2023, 5):
        is_tot = "TOTAL" in deb
        ws.cell(row=idx, column=1, value=deb).font = bold_font if is_tot else regular_font
        for c_idx, val in enumerate([c0, c30, c60, c90, gross], 2):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
        if lrate:
            c_r = ws.cell(row=idx, column=7, value=lrate)
            c_r.number_format = fmt_percent
            c_r.alignment = align_center
        else:
            ws.cell(row=idx, column=7, value="").alignment = align_center
        c_e = ws.cell(row=idx, column=8, value=ecl)
        c_n = ws.cell(row=idx, column=9, value=net)
        for c in [c_e, c_n]:
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
        for c in range(1, 10):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    # Trade and Other Payables Subledger
    ws.cell(row=12, column=1, value="TRADE AND OTHER PAYABLES SUBLEDGER (2022 - 2025)").font = section_font
    pay_headers = ["Creditor / Supplier / Accrual Account", "Category", "31 Dec 2022 (NGN)", "31 Dec 2023 (NGN)", "31 Dec 2024 (NGN)", "31 Dec 2025 (NGN)", "Payment Terms & Status"]
    style_header(ws, 13, pay_headers)
    
    payables_data = [
        ("Titan Piling Nig Ltd (Subcontractor)", "Trade Subcontractor", 0.00, 1200000.00, 2100000.00, 3400000.00, "30-day certificate settlement terms"),
        ("Prime Steel & Cement Ltd (Materials)", "Direct Material Supplier", 25000.00, 950000.00, 1750000.00, 3050000.00, "Commercial invoice credit facility"),
        ("PKF & Co. / Baker & Associates (Audit Fees)", "Audit Fee Accrual", 50000.00, 375000.00, 550000.00, 850000.00, "Statutory audit fee accrued annually"),
        ("Sundry Site Operating & Utilities Accruals", "Operational Accruals", 0.00, 245235.00, 425500.00, 683500.00, "Electricity, communication and site security"),
        ("TOTAL TRADE AND OTHER PAYABLES", "Current Liabilities", "=SUM(C14:C17)", "=SUM(D14:D17)", "=SUM(E14:E17)", "=SUM(F14:F17)", "TIED TO SFP NOTE 14"),
    ]
    
    for idx, (cred, cat, p22, p23, p24, p25, tms) in enumerate(payables_data, 14):
        is_tot = "TOTAL" in cred
        ws.cell(row=idx, column=1, value=cred).font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=2, value=cat).font = regular_font
        for c_idx, val in enumerate([p22, p23, p24, p25], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=7, value=tms).font = success_font if is_tot else italic_font
        ws.cell(row=idx, column=7).alignment = align_center if is_tot else align_left
        for c in range(1, 8):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

def build_sheet_19_related_parties_equity(wb):
    ws = wb.create_sheet(title="Related_Parties_Equity")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — RELATED PARTIES (IAS 24) & EQUITY ROLL-FORWARD").font = title_font
    ws.cell(row=2, column=1, value="Director Current Accounts, Trade Alias Movements & Statement of Changes in Equity").font = italic_font
    
    headers = ["Related Party / Entity Name", "Relationship & Nature", "31 Dec 2022 (NGN)", "31 Dec 2023 (NGN)", "31 Dec 2024 (NGN)", "31 Dec 2025 (NGN)", "Terms, Interest & Subordination Status"]
    style_header(ws, 4, headers)
    
    related_data = [
        ("Engr. Babatunde Goke", "Managing Director / Controlling Shareholder", 0.00, 4500000.00, 5800000.00, 7500000.00, "Unsecured, non-interest bearing long-term project funding facility; subordinated to third-party bank debt."),
        ("Baay Gokes (Business Name Alias)", "Trade Alias (Same Legal Entity RC 1526224)", 0.00, 0.00, 0.00, 0.00, "All bank accounts & transactions merged into Baay Projects Limited. Zero intercompany balance."),
        ("Baay Degok (Business Name Alias)", "Trade Alias (Same Legal Entity RC 1526224)", 0.00, 0.00, 0.00, 0.00, "All bank accounts & transactions merged into Baay Projects Limited. Zero intercompany balance."),
        ("TOTAL RELATED PARTY BORROWINGS", "Non-Current Liabilities", "=SUM(C5:C7)", "=SUM(D5:D7)", "=SUM(E5:E7)", "=SUM(F5:F7)", "TIED TO SFP NOTE 17"),
    ]
    
    for idx, (rp, rel, r22, r23, r24, r25, tms) in enumerate(related_data, 5):
        is_tot = "TOTAL" in rp
        ws.cell(row=idx, column=1, value=rp).font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=2, value=rel).font = regular_font
        for c_idx, val in enumerate([r22, r23, r24, r25], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=7, value=tms).font = success_font if is_tot else italic_font
        ws.cell(row=idx, column=7).alignment = align_center if is_tot else align_left
        for c in range(1, 8):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill

    # Statement of Changes in Equity Section
    ws.cell(row=11, column=1, value="STATEMENT OF CHANGES IN EQUITY (2022 - 2025)").font = section_font
    eq_headers = ["Equity Component", "Ordinary Share Capital (NGN)", "Retained Earnings (NGN)", "Total Shareholders' Equity (NGN)", "Audit Movement Verification"]
    style_header(ws, 12, eq_headers)
    
    eq_rows = [
        ("Balance at 1 January 2022", 1000000.00, 1650000.00, 2650000.00, "Opening Share Capital & Earnings"),
        ("Profit for the Year 2022", 0.00, 711708.00, 711708.00, "2022 Audited AFS Net Profit"),
        ("Balance at 31 December 2022", 1000000.00, 2361708.00, 3361708.00, "Audited 2022 Balance Sheet"),
        
        ("Profit for the Year 2023", 0.00, 5600235.00, 5600235.00, "Statement of Profit or Loss FY 2023"),
        ("Balance at 31 December 2023", 1000000.00, 7961943.00, 8961943.00, "Audited SFP 31 Dec 2023"),
        
        ("Profit for the Year 2024", 0.00, 9514775.00, 9514775.00, "Statement of Profit or Loss FY 2024"),
        ("Balance at 31 December 2024", 1000000.00, 17476718.00, 18476718.00, "Audited SFP 31 Dec 2024"),
        
        ("Profit for the Year 2025", 0.00, 14899870.00, 14899870.00, "Statement of Profit or Loss FY 2025"),
        ("Balance at 31 December 2025", 1000000.00, 32376588.00, 33376588.00, "Audited SFP 31 Dec 2025"),
    ]
    
    for idx, (comp, sc, re, tot, stat) in enumerate(eq_rows, 13):
        is_bal = "Balance at 31" in comp
        ws.cell(row=idx, column=1, value=comp).font = bold_font if is_bal else regular_font
        for c_idx, val in enumerate([sc, re, tot], 2):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_bal else regular_font
        c_st = ws.cell(row=idx, column=5, value=stat)
        c_st.font = success_font if is_bal else italic_font
        c_st.alignment = align_center
        for c in range(1, 6):
            ws.cell(row=idx, column=c).border = double_bottom if is_bal else thin_border
            if is_bal: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

print("sheet_builders_3.py completed with Sheets 15-19.")
