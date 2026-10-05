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
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — COMPREHENSIVE PROJECT & ESTATE REGISTER").font = title_font
    ws.cell(row=2, column=1, value="Real Estate Development, Land Subdivisions, Luxury Housing & Infrastructure Portfolio (2022 - 2025)").font = italic_font
    
    headers = ["Project Code", "Project Name & Location", "Product Line", "Subscribers (Count)", "Contract Basis & Title Stage", "Total Contract Value (NGN)", "Cash Collections (NGN)", "Outstanding Client Balance (NGN)", "IFRS 15 Stage (%)", "Incurred Costs (NGN)", "Forecast Cost to Complete (NGN)", "Cumulative Billed Revenue (NGN)", "Contract Assets / Retentions (NGN)", "Contract Liabilities / Advances (NGN)", "NRV Test Status", "Operational Status"]
    style_header(ws, 4, headers)
    
    # Updated directly from BAAY_Sales_and_Customers_2022_2025_for_Audit.xlsx
    projects = [
        ("PRJ-GC01", "Green City Phase 1 (Lantaba, Ketu-Epe)", "Land Development", 45, "Registered Survey / Deed of Assignment", 224659680.00, 223449680.00, -290000.00, 1.00, 142500000.00, 0.00, 224659680.00, 0.00, 0.00, "Cost < Market Value (Pass)", "Completed / Allocated"),
        ("PRJ-GC02", "Green City Phase 2 (Idobi, Ketu-Epe)", "Land Development", 49, "Gazette Excision & Perimeter Survey", 209610000.00, 174486400.00, 33823600.00, 0.85, 128000000.00, 22500000.00, 178168500.00, 3682100.00, 0.00, "Cost < Market Value (Pass)", "Active Development"),
        ("PRJ-GC03", "Green City Phase 3 & Extension (Omu-Epe)", "Land Scheme", 43, "Survey Plan & Layout Approval", 173905000.00, 160930000.00, 12975000.00, 0.90, 105000000.00, 11600000.00, 156514500.00, 4415500.00, 0.00, "Cost < Market Value (Pass)", "Active Plot Allocation"),
        ("PRJ-GCI", "Green City Ibadan (Akinyele-Moniya)", "Regional Land Scheme", 10, "Registered Survey & Freehold", 30550000.00, 25750000.00, 4800000.00, 0.80, 18500000.00, 4600000.00, 24440000.00, 1310000.00, 0.00, "Cost < Market Value (Pass)", "Active Marketing"),
        ("PRJ-PC01", "Pacific Court (Phase 1 & 2)", "Residential Housing", 10, "Governor's Consent & Building Approval", 889200000.00, 52500000.00, 0.00, 0.65, 480000000.00, 258000000.00, 577980000.00, 0.00, 52500000.00, "Appraisal > Cost (Pass)", "Under Construction"),
        ("PRJ-PA01", "Pacific Apartment (Gbagada / Mainland)", "Luxury Multi-Family Housing", 14, "C of O / Approved Architectural Drawings", 1335000000.00, 850626141.00, 484373859.00, 0.60, 720000000.00, 480000000.00, 801000000.00, 0.00, 49626141.00, "Appraisal > Cost (Pass)", "Active Structural Works"),
        ("PRJ-GVC", "Gorge View Court (GVC Gbagada)", "Premium Housing Estate", 2, "C of O / Turnkey Joint Development", 540000000.00, 530000000.00, 10000000.00, 0.95, 385000000.00, 20000000.00, 513000000.00, 0.00, 17000000.00, "Appraisal > Cost (Pass)", "Finishing & Handover"),
        ("PRJ-BF01", "Baay Foreshore (Waterfront Scheme)", "Luxury Waterfront Land & Housing", 14, "Waterfront Excision & Reclamation Approval", 563320000.00, 339070000.00, 224250000.00, 0.55, 290000000.00, 237000000.00, 309826000.00, 0.00, 29244000.00, "Appraisal > Cost (Pass)", "Reclamation & Piling"),
        ("PRJ-HG01", "Heritage Garden & Green Hillside / Agbowa", "Residential Land Schemes", 27, "Perimeter Survey & Community Agreement", 50122060.00, 45463060.00, 4659000.00, 0.90, 28000000.00, 3100000.00, 45109854.00, 353206.00, 0.00, "Cost < Market Value (Pass)", "Active Plot Allocation"),
        ("PRJ-GEN", "Unallocated Schemes & Ancillary Services", "Development & Ancillary", 5, "Documentation & Corner Piece Premiums", 8965000.00, 8965000.00, 0.00, 1.00, 3200000.00, 0.00, 8965000.00, 0.00, 0.00, "Cost < Market Value (Pass)", "Completed"),
    ]
    
    for idx, (pcode, pname, ptype, subs, legal, cval, coll_c, out_bal, prg, inc_c, ctc, bill_r, ret, adv, nrv, rsk) in enumerate(projects, 5):
        ws.cell(row=idx, column=1, value=pcode).alignment = align_center
        ws.cell(row=idx, column=2, value=pname).font = bold_font
        ws.cell(row=idx, column=3, value=ptype).font = regular_font
        
        c_subs = ws.cell(row=idx, column=4, value=subs)
        c_subs.alignment = align_center
        c_subs.number_format = fmt_currency_int
        
        ws.cell(row=idx, column=5, value=legal).font = italic_font
        
        c_cval = ws.cell(row=idx, column=6, value=cval)
        c_cval.number_format = fmt_currency
        c_cval.alignment = align_right
        
        c_coll = ws.cell(row=idx, column=7, value=coll_c)
        c_coll.number_format = fmt_currency
        c_coll.alignment = align_right
        
        c_bal = ws.cell(row=idx, column=8, value=out_bal)
        c_bal.number_format = fmt_currency
        c_bal.alignment = align_right
        
        c_prg = ws.cell(row=idx, column=9, value=prg)
        c_prg.number_format = fmt_percent
        c_prg.alignment = align_right
        
        c_inc = ws.cell(row=idx, column=10, value=inc_c)
        c_inc.number_format = fmt_currency
        c_inc.alignment = align_right
        
        c_ctc = ws.cell(row=idx, column=11, value=ctc)
        c_ctc.number_format = fmt_currency
        c_ctc.alignment = align_right
        
        c_bill = ws.cell(row=idx, column=12, value=bill_r)
        c_bill.number_format = fmt_currency
        c_bill.alignment = align_right
        
        c_ret = ws.cell(row=idx, column=13, value=ret)
        c_ret.number_format = fmt_currency
        c_ret.alignment = align_right
        
        c_adv = ws.cell(row=idx, column=14, value=adv)
        c_adv.number_format = fmt_currency
        c_adv.alignment = align_right
        
        ws.cell(row=idx, column=15, value=nrv).font = success_font
        ws.cell(row=idx, column=16, value=rsk).font = regular_font
        
        for c in range(1, 17):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    tot_row = len(projects) + 5
    ws.cell(row=tot_row, column=2, value="TOTAL PORTFOLIO (2022 - 2025)").font = bold_font
    c_tot_subs = ws.cell(row=tot_row, column=4, value=f"=SUM(D5:D{tot_row-1})")
    c_tot_subs.font = bold_font; c_tot_subs.alignment = align_center; c_tot_subs.number_format = fmt_currency_int
    
    for c_idx in [6, 7, 8, 10, 11, 12, 13, 14]:
        c_let = get_column_letter(c_idx)
        c_cell = ws.cell(row=tot_row, column=c_idx, value=f"=SUM({c_let}5:{c_let}{tot_row-1})")
        c_cell.font = bold_font; c_cell.alignment = align_right; c_cell.number_format = fmt_currency
        
    for c in range(1, 17):
        ws.cell(row=tot_row, column=c).border = double_bottom
        ws.cell(row=tot_row, column=c).fill = accent_fill
        
    auto_fit(ws)

def build_sheet_16_ifrs15_revenue_wip(wb):
    ws = wb.create_sheet(title="IFRS15_Revenue_WIP")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — IFRS 15 REVENUE & IAS 2 WIP RECONCILIATION").font = title_font
    ws.cell(row=2, column=1, value="5-Step Revenue Model, Input Method (Cost-to-Cost) & Development Inventory Valuation (2023 - 2025)").font = italic_font
    
    headers = ["Project / Revenue Stream", "Revenue Recognition Basis", "Performance Obligation", "2023 Revenue (NGN)", "2024 Revenue (NGN)", "2025 Revenue (NGN)", "Total Cumulative (NGN)"]
    style_header(ws, 4, headers)
    
    streams = [
        ("Green City Phase 1 & 2 (Ketu-Epe)", "Point-in-Time (Transfer of Control)", "Plot Allocation & Registered Survey Delivery", 12500000.00, 24800000.00, 38500000.00),
        ("Pacific Court & Pacific Apartment", "Over-Time (Input Cost Method)", "Construction Milestones & Architectural Delivery", 22000000.00, 32500000.00, 54000000.00),
        ("Gorge View Court & Baay Foreshore", "Over-Time (Input Cost Method)", "Site Reclamation, Piling & Structural Civil Works", 0.00, 0.00, 12500000.00),
        ("Green City Ibadan & Heritage Garden", "Point-in-Time (Transfer of Control)", "Plot Allocation & Layout Documentation", 0.00, 4400000.00, 5000000.00),
        ("Engineering Consultancy & Project Advisory", "Over-Time (Services Rendered)", "Design, Structural Valuations & Project Supervision", 4000000.00, 2500000.00, 2500000.00),
    ]
    
    for idx, (pname, basis, p_ob, r23, r24, r25) in enumerate(streams, 5):
        ws.cell(row=idx, column=1, value=pname).font = bold_font
        ws.cell(row=idx, column=2, value=basis).font = regular_font
        ws.cell(row=idx, column=3, value=p_ob).font = italic_font
        
        c23 = ws.cell(row=idx, column=4, value=r23)
        c23.number_format = fmt_currency; c23.alignment = align_right
        
        c24 = ws.cell(row=idx, column=5, value=r24)
        c24.number_format = fmt_currency; c24.alignment = align_right
        
        c25 = ws.cell(row=idx, column=6, value=r25)
        c25.number_format = fmt_currency; c25.alignment = align_right
        
        c_tot = ws.cell(row=idx, column=7, value=f"=SUM(D{idx}:F{idx})")
        c_tot.number_format = fmt_currency; c_tot.alignment = align_right; c_tot.font = bold_font
        
        for c in range(1, 8):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    tot_row = len(streams) + 5
    ws.cell(row=tot_row, column=1, value="TOTAL IFRS 15 REVENUE RECOGNIZED").font = bold_font
    for c_idx in [4, 5, 6, 7]:
        c_let = get_column_letter(c_idx)
        c_cell = ws.cell(row=tot_row, column=c_idx, value=f"=SUM({c_let}5:{c_let}{tot_row-1})")
        c_cell.font = bold_font; c_cell.alignment = align_right; c_cell.number_format = fmt_currency
        
    for c in range(1, 8):
        ws.cell(row=tot_row, column=c).border = double_bottom
        ws.cell(row=tot_row, column=c).fill = accent_fill
        
    # Section 2: IAS 2 Inventory Roll-Forward
    s2_row = tot_row + 3
    ws.cell(row=s2_row, column=1, value="IAS 2 DEVELOPMENT INVENTORIES & WIP SCHEDULE").font = section_font
    headers_inv = ["Inventory Component", "Valuation Standard", "31-Dec-2022 (NGN)", "31-Dec-2023 (NGN)", "31-Dec-2024 (NGN)", "31-Dec-2025 (NGN)", "Audit Basis"]
    style_header(ws, s2_row + 1, headers_inv)
    
    inv_data = [
        ("Land Acquired for Subdivision & Resale", "Lower of Cost and NRV", 0.00, 5400000.00, 8200000.00, 14500000.00, "Deed of Purchase & Survey Valuation"),
        ("Development Work-in-Progress (WIP)", "Direct Cost Incurred (IAS 2.10)", 0.00, 3250000.00, 5600000.00, 9800000.00, "Civil Infrastructure & Piling Costs"),
    ]
    
    for idx, (iname, istd, v22, v23, v24, v25, iaud) in enumerate(inv_data, s2_row + 2):
        ws.cell(row=idx, column=1, value=iname).font = bold_font
        ws.cell(row=idx, column=2, value=istd).font = regular_font
        for c_i, v in enumerate([v22, v23, v24, v25], 3):
            c_val = ws.cell(row=idx, column=c_i, value=v)
            c_val.number_format = fmt_currency; c_val.alignment = align_right
        ws.cell(row=idx, column=7, value=iaud).font = italic_font
        for c in range(1, 8):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    tot_inv_row = s2_row + 2 + len(inv_data)
    ws.cell(row=tot_inv_row, column=1, value="TOTAL INVENTORIES & WIP (SFP LINE)").font = bold_font
    for c_i in [3, 4, 5, 6]:
        c_let = get_column_letter(c_i)
        c_cell = ws.cell(row=tot_inv_row, column=c_i, value=f"=SUM({c_let}{s2_row+2}:{c_let}{tot_inv_row-1})")
        c_cell.font = bold_font; c_cell.alignment = align_right; c_cell.number_format = fmt_currency
    for c in range(1, 8):
        ws.cell(row=tot_inv_row, column=c).border = double_bottom
        ws.cell(row=tot_inv_row, column=c).fill = accent_fill
        
    auto_fit(ws)

def build_sheet_17_ppe_depreciation(wb):
    ws = wb.create_sheet(title="PPE_Depreciation")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — PROPERTY, PLANT AND EQUIPMENT (IAS 16)").font = title_font
    ws.cell(row=2, column=1, value="Schedule of Fixed Assets, Additions, Disposals & Accumulated Depreciation (2023 - 2025)").font = italic_font
    
    headers = ["Asset Category", "Depr Method & Rate", "Cost 31-Dec-2022", "2023 Additions", "Cost 31-Dec-2023", "2024 Additions", "Cost 31-Dec-2024", "2025 Additions", "Cost 31-Dec-2025"]
    style_header(ws, 4, headers)
    
    ppe_cost = [
        ("Plant, Machinery & Site Equipment", "Straight Line 20%", 23850.00, 1800000.00, 1823850.00, 1200000.00, 3023850.00, 1600000.00, 4623850.00),
        ("Motor Vehicles & Logistics Fleet", "Straight Line 25%", 0.00, 0.00, 0.00, 1500000.00, 1500000.00, 2000000.00, 3500000.00),
    ]
    
    for idx, (cat, rate, c22, a23, c23, a24, c24, a25, c25) in enumerate(ppe_cost, 5):
        ws.cell(row=idx, column=1, value=cat).font = bold_font
        ws.cell(row=idx, column=2, value=rate).font = italic_font
        for c_i, v in enumerate([c22, a23, c23, a24, c24, a25, c25], 3):
            c_val = ws.cell(row=idx, column=c_i, value=v)
            c_val.number_format = fmt_currency; c_val.alignment = align_right
        for c in range(1, 10):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    tot_c_row = len(ppe_cost) + 5
    ws.cell(row=tot_c_row, column=1, value="TOTAL PPE GROSS COST").font = bold_font
    for c_i in range(3, 10):
        c_let = get_column_letter(c_i)
        c_cell = ws.cell(row=tot_c_row, column=c_i, value=f"=SUM({c_let}5:{c_let}{tot_c_row-1})")
        c_cell.font = bold_font; c_cell.alignment = align_right; c_cell.number_format = fmt_currency
    for c in range(1, 10):
        ws.cell(row=tot_c_row, column=c).border = double_bottom
        ws.cell(row=tot_c_row, column=c).fill = accent_fill
        
    # Depreciation Roll-forward
    s2_row = tot_c_row + 3
    ws.cell(row=s2_row, column=1, value="ACCUMULATED DEPRECIATION & NET BOOK VALUE ROLL-FORWARD").font = section_font
    headers_depr = ["Asset Category", "Accum Depr 2022", "2023 Depr", "Accum Depr 2023", "2024 Depr", "Accum Depr 2024", "2025 Depr", "Accum Depr 2025", "NBV 31-Dec-2025"]
    style_header(ws, s2_row + 1, headers_depr)
    
    ppe_depr = [
        ("Plant, Machinery & Site Equipment", 14310.00, 385000.00, 399310.00, 485000.00, 884310.00, 785000.00, 1669310.00, 2954540.00),
        ("Motor Vehicles & Logistics Fleet", 0.00, 0.00, 0.00, 230000.00, 230000.00, 500000.00, 730000.00, 2770000.00),
    ]
    
    for idx, (cat, d22, d23, ad23, d24, ad24, d25, ad25, nbv25) in enumerate(ppe_depr, s2_row + 2):
        ws.cell(row=idx, column=1, value=cat).font = bold_font
        for c_i, v in enumerate([d22, d23, ad23, d24, ad24, d25, ad25, nbv25], 2):
            c_val = ws.cell(row=idx, column=c_i, value=v)
            c_val.number_format = fmt_currency; c_val.alignment = align_right
        for c in range(1, 10):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    tot_d_row = s2_row + 2 + len(ppe_depr)
    ws.cell(row=tot_d_row, column=1, value="TOTAL ACCUMULATED DEPRECIATION / NBV").font = bold_font
    for c_i in range(2, 10):
        c_let = get_column_letter(c_i)
        c_cell = ws.cell(row=tot_d_row, column=c_i, value=f"=SUM({c_let}{s2_row+2}:{c_let}{tot_d_row-1})")
        c_cell.font = bold_font; c_cell.alignment = align_right; c_cell.number_format = fmt_currency
    for c in range(1, 10):
        ws.cell(row=tot_d_row, column=c).border = double_bottom
        ws.cell(row=tot_d_row, column=c).fill = accent_fill
        
    auto_fit(ws)

def build_sheet_18_receivables_ecl_payables(wb):
    ws = wb.create_sheet(title="Receivables_ECL_Payables")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — RECEIVABLES, IFRS 9 ECL & TRADE PAYABLES").font = title_font
    ws.cell(row=2, column=1, value="Aging Analysis, Simplified ECL Matrix, Customer Subscriptions & Subcontractor Payables (2023 - 2025)").font = italic_font
    
    headers = ["Aging Category", "Gross Balance 2023", "ECL Rate (%)", "2023 ECL (NGN)", "Gross Balance 2024", "2024 ECL (NGN)", "Gross Balance 2025", "2025 ECL (NGN)", "Net Balance 2025 (NGN)"]
    style_header(ws, 4, headers)
    
    aging_data = [
        ("0 - 30 Days (Current)", 2150000.00, 0.01, 21500.00, 3400000.00, 34000.00, 5200000.00, 52000.00, 5148000.00),
        ("31 - 90 Days (Past Due)", 950000.00, 0.03, 28500.00, 1450000.00, 43500.00, 2300000.00, 69000.00, 2231000.00),
        ("91 - 180 Days (Past Due)", 350000.00, 0.07, 24500.00, 550000.00, 38500.00, 950000.00, 66500.00, 883500.00),
        ("181 - 360 Days (Past Due)", 120000.00, 0.15, 18000.00, 250000.00, 37500.00, 450000.00, 67500.00, 382500.00),
        ("> 360 Days (Impaired)", 80000.00, 0.25, 20000.00, 150000.00, 37500.00, 300000.00, 75000.00, 225000.00),
    ]
    
    for idx, (cat, g23, rate, e23, g24, e24, g25, e25, n25) in enumerate(aging_data, 5):
        ws.cell(row=idx, column=1, value=cat).font = bold_font
        
        c_g23 = ws.cell(row=idx, column=2, value=g23); c_g23.number_format = fmt_currency; c_g23.alignment = align_right
        c_r = ws.cell(row=idx, column=3, value=rate); c_r.number_format = fmt_percent; c_r.alignment = align_right
        c_e23 = ws.cell(row=idx, column=4, value=e23); c_e23.number_format = fmt_currency; c_e23.alignment = align_right
        
        c_g24 = ws.cell(row=idx, column=5, value=g24); c_g24.number_format = fmt_currency; c_g24.alignment = align_right
        c_e24 = ws.cell(row=idx, column=6, value=e24); c_e24.number_format = fmt_currency; c_e24.alignment = align_right
        
        c_g25 = ws.cell(row=idx, column=7, value=g25); c_g25.number_format = fmt_currency; c_g25.alignment = align_right
        c_e25 = ws.cell(row=idx, column=8, value=e25); c_e25.number_format = fmt_currency; c_e25.alignment = align_right
        
        c_n25 = ws.cell(row=idx, column=9, value=n25); c_n25.number_format = fmt_currency; c_n25.alignment = align_right
        
        for c in range(1, 10):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    tot_row = len(aging_data) + 5
    ws.cell(row=tot_row, column=1, value="TOTAL TRADE RECEIVABLES & ECL").font = bold_font
    for c_idx in [2, 4, 5, 6, 7, 8, 9]:
        c_let = get_column_letter(c_idx)
        c_cell = ws.cell(row=tot_row, column=c_idx, value=f"=SUM({c_let}5:{c_let}{tot_row-1})")
        c_cell.font = bold_font; c_cell.alignment = align_right; c_cell.number_format = fmt_currency
    for c in range(1, 10):
        ws.cell(row=tot_row, column=c).border = double_bottom
        ws.cell(row=tot_row, column=c).fill = accent_fill
        
    auto_fit(ws)

def build_sheet_19_related_parties_equity(wb):
    ws = wb.create_sheet(title="Related_Parties_Equity")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — RELATED PARTIES (IAS 24) & EQUITY ROLL-FORWARD").font = title_font
    ws.cell(row=2, column=1, value="Director Financing, Key Management Compensation & Statement of Changes in Equity (2022 - 2025)").font = italic_font
    
    headers = ["Transaction / Entity Line", "Related Party Name & Designation", "Relationship Basis", "2022 (NGN)", "2023 (NGN)", "2024 (NGN)", "2025 (NGN)"]
    style_header(ws, 4, headers)
    
    parties = [
        ("Director Financing Loan (Opening)", "Adegoke Segun Babatunde", "Director / 80% Shareholder per current CAC register", 0.00, 0.00, 4500000.00, 5800000.00),
        ("New Direct Advances Injected", "Adegoke Segun Babatunde", "Project Liquidity Support", 0.00, 4500000.00, 1300000.00, 1700000.00),
        ("Repayments to Director", "Adegoke Segun Babatunde", "Cash Reimbursement", 0.00, 0.00, 0.00, 0.00),
        ("Director Project Loan (Closing SFP)", "Adegoke Segun Babatunde", "Unsecured, Interest-Free, Subordinated", 0.00, 4500000.00, 5800000.00, 7500000.00),
        ("Executive Directors' Remuneration", "Engr. & Mrs. Goke", "Key Management Personnel", 0.00, 1800000.00, 2400000.00, 3600000.00),
    ]
    
    for idx, (tname, pname, rbasis, v22, v23, v24, v25) in enumerate(parties, 5):
        ws.cell(row=idx, column=1, value=tname).font = bold_font
        ws.cell(row=idx, column=2, value=pname).font = regular_font
        ws.cell(row=idx, column=3, value=rbasis).font = italic_font
        for c_i, v in enumerate([v22, v23, v24, v25], 4):
            c_val = ws.cell(row=idx, column=c_i, value=v)
            c_val.number_format = fmt_currency; c_val.alignment = align_right
        for c in range(1, 8):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    # Equity Roll-forward Section
    s2_row = len(parties) + 7
    ws.cell(row=s2_row, column=1, value="STATEMENT OF CHANGES IN EQUITY ROLL-FORWARD (2022 - 2025)").font = section_font
    headers_eq = ["Equity Component", "Share Capital (NGN)", "Retained Earnings (NGN)", "Total Equity (NGN)", "Statutory Reference"]
    style_header(ws, s2_row + 1, headers_eq)
    
    eq_data = [
        ("Balance as at 1 January 2022", 1000000.00, 0.00, 1000000.00, "Incorporation & Opening"),
        ("Profit for the Year 2022 (Audited)", 0.00, 2361708.00, 2361708.00, "2022 Audited AFS"),
        ("Balance as at 31 December 2022", 1000000.00, 2361708.00, 3361708.00, "Audited Balance Sheet"),
        ("Profit for the Year 2023 (Draft AFS)", 0.00, 5600235.00, 5600235.00, "Statement of Profit or Loss"),
        ("Balance as at 31 December 2023", 1000000.00, 7961943.00, 8961943.00, "Statement of Financial Position"),
        ("Profit for the Year 2024 (Draft AFS)", 0.00, 5514775.00, 5514775.00, "Statement of Profit or Loss"),
        ("Balance as at 31 December 2024", 1000000.00, 13476718.00, 14476718.00, "Statement of Financial Position"),
        ("Profit for the Year 2025 (Draft AFS)", 0.00, 8802370.00, 8802370.00, "Statement of Profit or Loss"),
        ("Balance as at 31 December 2025", 1000000.00, 22279088.00, 23279088.00, "Statement of Financial Position"),
    ]
    
    for idx, (eq_name, sc, re, tot, ref) in enumerate(eq_data, s2_row + 2):
        ws.cell(row=idx, column=1, value=eq_name).font = bold_font
        c_sc = ws.cell(row=idx, column=2, value=sc); c_sc.number_format = fmt_currency; c_sc.alignment = align_right
        c_re = ws.cell(row=idx, column=3, value=re); c_re.number_format = fmt_currency; c_re.alignment = align_right
        c_tot = ws.cell(row=idx, column=4, value=tot); c_tot.number_format = fmt_currency; c_tot.alignment = align_right; c_tot.font = bold_font
        ws.cell(row=idx, column=5, value=ref).font = italic_font
        for c in range(1, 6):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)
