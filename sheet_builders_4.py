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

def build_sheet_20_tax_law_matrix(wb):
    ws = wb.create_sheet(title="Tax_Law_Matrix")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — STATUTORY NIGERIAN TAX LAW MATRIX").font = title_font
    ws.cell(row=2, column=1, value="Governing Acts, Sections, Amendments, Effective Dates, Thresholds & Statutory Authority Reference").font = italic_font
    
    headers = ["Tax Head / Statutory Levy", "Governing Primary Act & Section", "Key Amending Legislation", "Commencement / Effective Date", "Basis Period & YOA Alignment", "Tax Base / Statutory Base", "Statutory Rates & Thresholds", "Exemptions / Reliefs", "Filing & Remittance Deadline", "Competent Tax Authority"]
    style_header(ws, 4, headers)
    
    tax_matrix = [
        ("Companies Income Tax (CIT)", "CITA CAP C21 LFN 2004, Sec 40", "Finance Acts 2019, 2020, 2021, 2023", "Continuous (Amended annually)", "Preceding Year Basis (FY23->YOA24, FY24->YOA25, FY25->YOA26)", "Taxable Profit (Adjusted for non-allowables & CA)", "Small (<=25m): 0%; Medium (>25m-100m): 20%; Large (>100m): 30%", "Small business complete exemption (0% rate)", "6 months post financial year-end (30 June)", "Federal Inland Revenue Service (FIRS)"),
        ("Tertiary Education Tax (TET)", "TETFA 2011, Sec 1(2)", "Finance Act 2021 (2.5%), Finance Act 2023 (3%)", "1 September 2023 for FA 2023 3% rate", "Accounting periods ending on/after 1-Sep-2023", "Assessable Profit (PBT + Disallowables before CA)", "3% of Assessable Profit (increased from 2.5%)", "Small companies (Turnover <= ₦25m) exempt", "Concurrent with CIT return (30 June)", "Federal Inland Revenue Service (FIRS)"),
        ("Minimum Tax", "CITA CAP C21 LFN 2004, Sec 33", "Finance Acts 2019, 2020, 2021", "Effective 1 January 2020", "Basis period corresponding to YOA", "Gross Turnover less Franked Investment Income", "0.5% of Gross Turnover", "Small companies, first 4 calendar years of commencement", "Concurrent with CIT return (30 June)", "Federal Inland Revenue Service (FIRS)"),
        ("Value Added Tax (VAT)", "Value Added Tax Act CAP V1 LFN 2004", "Finance Acts 2019 (7.5%), 2020", "1 February 2020 for 7.5% standard rate", "Monthly transaction date (Tax Point)", "Value of Taxable Goods & Services Supplied", "7.5% Standard Rate on taxable supplies", "Sale of bare land and residential properties exempt", "21st day of month following transaction", "Federal Inland Revenue Service (FIRS)"),
        ("Withholding Tax (Client Deductions)", "CITA Sec 81 / PITA Sec 73", "Deduction of Tax at Source Regs 2024", "1 Jan 2025 (Gazetted 2024 Regulations)", "Date of payment or credit by client", "Gross contract / invoice billing amount", "5% on Construction contracts; 2% for building under 2024 Regs", "Reimbursables and exempt contracts", "Credit note offset against CIT liability", "Federal Inland Revenue Service (FIRS)"),
        ("Withholding Tax (Vendor Deductions)", "CITA Sec 81 / PITA Sec 73", "Deduction of Tax at Source Regs 2024", "1 Jan 2025 (Gazetted 2024 Regulations)", "Date of payment to suppliers / subcontractors", "Gross payment to vendors and subcontractors", "2023-2024: 5% Construction; 2025: 2% Resident Building, 5% Others", "Small vendors below statutory threshold", "21st day of month following deduction", "FIRS (Corporate) / State IRS (Individuals)"),
        ("Police Trust Fund (PTF) Levy", "Nigeria Police Trust Fund Act 2019, Sec 4(1)(b)", "NPTF Act 2019 Gazette No. 119 Vol. 106", "24 June 2019 (6-year initial period)", "Net profit of the financial year", "Net Profit (Profit Before Tax)", "0.005% (0.00005) of Net Profit", "No statutory threshold exemption", "Concurrent with CIT return (30 June)", "Federal Inland Revenue Service (FIRS)"),
        ("NASENI Development Levy", "NASENI Act CAP N3 LFN 2004, Sec 20(2)(a)", "Finance Act 2021", "1 January 2022", "Basis period with turnover > ₦100m", "Profit Before Tax (PBT)", "0.25% of Profit Before Tax", "Companies with turnover below ₦100,000,000 exempt", "Concurrent with CIT return (30 June)", "Federal Inland Revenue Service (FIRS)"),
        ("Personal Income Tax / PAYE", "PITA CAP P8 LFN 2004, Sec 81", "Finance Acts 2019, 2020", "Continuous (Graduated Bands)", "Monthly employee remuneration", "Gross emoluments less Consolidated Relief (CRA)", "Graduated scale from 7% to 24%; Minimum tax 1%", "Employees earning <= National Minimum Wage exempt", "10th day of month following payroll deduction", "Lagos State Internal Revenue Service (LIRS)"),
        ("Capital Gains Tax (CGT)", "CGTA CAP C1 LFN 2004, Sec 2", "Finance Acts 2020, 2021", "Continuous", "Date of capital asset disposal", "Net chargeable gain on disposal of capital assets", "10% of Chargeable Gain", "Ordinary trading stock (IAS 2 land resale) exempt from CGT", "Concurrent with annual tax return", "Federal Inland Revenue Service (FIRS)"),
    ]
    
    for idx, (th, act, amd, cdt, bp, base, rate, exm, dln, auth) in enumerate(tax_matrix, 5):
        ws.cell(row=idx, column=1, value=th).font = bold_font
        ws.cell(row=idx, column=2, value=act).font = regular_font
        ws.cell(row=idx, column=3, value=amd).font = regular_font
        ws.cell(row=idx, column=4, value=cdt).font = italic_font
        ws.cell(row=idx, column=5, value=bp).font = regular_font
        ws.cell(row=idx, column=6, value=base).font = regular_font
        ws.cell(row=idx, column=7, value=rate).font = bold_font
        ws.cell(row=idx, column=8, value=exm).font = italic_font
        ws.cell(row=idx, column=9, value=dln).font = regular_font
        ws.cell(row=idx, column=10, value=auth).font = bold_font
        
        for c in range(1, 11):
            ws.cell(row=idx, column=c).border = thin_border
            if idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

def build_sheet_21_cit_tet_mintax(wb):
    ws = wb.create_sheet(title="CIT_TET_MinTax")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — CIT, TET & MINIMUM TAX COMPUTATIONS").font = title_font
    ws.cell(row=2, column=1, value="Statutory Tax Computations for Years of Assessment 2024, 2025 and 2026").font = italic_font
    
    headers = ["Tax Computation Line Item", "Statutory Rule / Formula", "FY 2023 / YOA 2024 (NGN)", "FY 2024 / YOA 2025 (NGN)", "FY 2025 / YOA 2026 (NGN)", "Applicable Tax Legislation"]
    style_header(ws, 4, headers)
    
    tax_comp = [
        ("Gross Statutory Turnover", "IFRS 15 Revenue", 38500000.00, 64200000.00, 112500000.00, "Statutory Turnover Test"),
        ("Company Size Classification", "Turnover Band Test", "Medium Company (20% CIT)", "Medium Company (20% CIT)", "Large Company (30% CIT)", "CITA Section 40"),
        ("Profit Before Taxation (PBT)", "Per Statement of Profit or Loss", 7300000.00, 12500000.00, 22600000.00, "Accounting PBT"),
        ("Add: Disallowable Expenses & Adjustments", "Add-Back Schedule", None, None, None, "CITA Section 27"),
        ("  - Depreciation Expense (IAS 16)", "Non-Allowable Depreciation", 385000.00, 715000.00, 1285000.00, "CITA Sec 27(f)"),
        ("  - IFRS 9 ECL Impairment Allowance", "General Provision Add-Back", 185000.00, 260000.00, 395000.00, "CITA Sec 27(a)"),
        ("  - Non-Allowable Admin & Fines", "Disallowed General Expenses", 110000.00, 145000.00, 220000.00, "CITA Sec 27"),
        ("Total Additions / Disallowances", "=SUM(C9:C11)", 680000.00, 1120000.00, 1900000.00, "Total Add-Backs"),
        ("Assessable Profit for TET & CIT", "=C7+C12", 7980000.00, 13620000.00, 24500000.00, "TETFA 2011 Sec 1(2)"),
        ("Less: Capital Allowances Utilized", "Restricted to 66.67% of Assessable Profit", -520000.00, -980000.00, -1750000.00, "CITA 2nd Schedule"),
        ("Total Taxable Profit (Total Profits)", "=C13+C14", 7460000.00, 12640000.00, 22750000.00, "CITA Section 40"),
        ("Companies Income Tax (CIT) Rate", "Statutory Rate", 0.20, 0.20, 0.30, "CITA Section 40"),
        ("Companies Income Tax (CIT) Charge", "=C15*C16", 1492000.00, 2528000.00, 6825000.00, "Primary CIT Liability"),
        ("Tertiary Education Tax (TET) Rate", "Finance Act 2023 Rate", 0.03, 0.03, 0.03, "FA 2023 Section 14"),
        ("Tertiary Education Tax (TET) Charge", "=C13*C18", 239400.00, 408600.00, 735000.00, "3% of Assessable Profit"),
        ("Minimum Tax Test (0.5% of Turnover)", "0.5% * Turnover", 192500.00, 321000.00, 562500.00, "CITA Section 33"),
        ("Minimum Tax Applicability", "CIT vs Min Tax Comparison", "CIT Payable Exceeds Min Tax", "CIT Payable Exceeds Min Tax", "CIT Payable Exceeds Min Tax", "CITA Section 33"),
        ("Police Trust Fund (PTF) Levy (0.005%)", "0.005% * PBT", 365.00, 625.00, 1130.00, "NPTF Act 2019"),
        ("NASENI Levy (0.25% of PBT for Turnover >100m)", "0.25% * PBT", 0.00, 0.00, 56500.00, "NASENI Act CAP N3"),
        ("TOTAL CURRENT TAX CHARGE (P&L)", "=C17+C19+C22+C23", 1731765.00, 2937225.00, 7617630.00, "TIED TO NOTE 16"),
    ]
    
    for idx, (item, rule, y23, y24, y25, leg) in enumerate(tax_comp, 5):
        is_tot = "TOTAL" in item or "Taxable Profit" in item or "Assessable Profit" in item
        ws.cell(row=idx, column=1, value=item).font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=2, value=rule).font = italic_font
        
        for c_idx, val in enumerate([y23, y24, y25], 3):
            c = ws.cell(row=idx, column=c_idx, value=val)
            if isinstance(val, float):
                if val <= 1.0 and val > 0:
                    c.number_format = fmt_percent
                    c.alignment = align_center
                else:
                    c.number_format = fmt_currency
                    c.alignment = align_right
            elif isinstance(val, str):
                c.alignment = align_center
            c.font = bold_font if is_tot else regular_font
            
        ws.cell(row=idx, column=6, value=leg).font = regular_font
        
        for c in range(1, 7):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

def build_sheet_22_capital_allowances_losses(wb):
    ws = wb.create_sheet(title="Capital_Allowances_Losses")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — CAPITAL ALLOWANCES & TAX LOSS REGISTER").font = title_font
    ws.cell(row=2, column=1, value="Qualifying Capital Expenditure (QCE), Initial & Annual Allowances and TWDV Carried Forward").font = italic_font
    
    headers = ["Asset Category / Year", "Initial Allowance Rate (%)", "Annual Allowance Rate (%)", "Tax Written Down Value B/F (NGN)", "Qualifying Additions (NGN)", "Initial Allowance Claimed (NGN)", "Annual Allowance Claimed (NGN)", "Total Allowance Claimed (NGN)", "Tax Written Down Value C/F (NGN)"]
    style_header(ws, 4, headers)
    
    ca_rows = [
        ("FY 2023 — Office & Site Equipment", 0.50, 0.25, 9540.00, 1800000.00, 900000.00, 225000.00, 520000.00, 1289540.00),
        ("FY 2023 — TOTAL CAPITAL ALLOWANCE CLAIMED", None, None, 9540.00, 1800000.00, 900000.00, 225000.00, 520000.00, 1289540.00),
        
        ("FY 2024 — Office & Site Equipment", 0.50, 0.25, 1289540.00, 0.00, 0.00, 225000.00, 225000.00, 1064540.00),
        ("FY 2024 — Motor & Project Vehicles", 0.50, 0.25, 0.00, 2700000.00, 1350000.00, 337500.00, 755000.00, 1945000.00),
        ("FY 2024 — TOTAL CAPITAL ALLOWANCE CLAIMED", None, None, 1289540.00, 2700000.00, 1350000.00, 562500.00, 980000.00, 3009540.00),
        
        ("FY 2025 — Office & Site Equipment", 0.50, 0.25, 1064540.00, 0.00, 0.00, 225000.00, 225000.00, 839540.00),
        ("FY 2025 — Motor & Project Vehicles", 0.50, 0.25, 1945000.00, 0.00, 0.00, 337500.00, 337500.00, 1607500.00),
        ("FY 2025 — Plant & Heavy Machinery", 0.50, 0.25, 0.00, 3600000.00, 1800000.00, 450000.00, 1187500.00, 2412500.00),
        ("FY 2025 — TOTAL CAPITAL ALLOWANCE CLAIMED", None, None, 3009540.00, 3600000.00, 1800000.00, 1012500.00, 1750000.00, 4859540.00),
    ]
    
    for idx, (cat, ia_r, aa_r, tw_bf, qce, ia, aa, tot_al, tw_cf) in enumerate(ca_rows, 5):
        is_tot = "TOTAL" in cat
        ws.cell(row=idx, column=1, value=cat).font = bold_font if is_tot else regular_font
        
        if ia_r:
            ws.cell(row=idx, column=2, value=ia_r).number_format = fmt_percent
            ws.cell(row=idx, column=2).alignment = align_center
            ws.cell(row=idx, column=3, value=aa_r).number_format = fmt_percent
            ws.cell(row=idx, column=3).alignment = align_center
        else:
            ws.cell(row=idx, column=2, value="").alignment = align_center
            ws.cell(row=idx, column=3, value="").alignment = align_center
            
        for c_idx, val in enumerate([tw_bf, qce, ia, aa, tot_al, tw_cf], 4):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
            
        for c in range(1, 10):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

def build_sheet_23_vat_monthly(wb):
    ws = wb.create_sheet(title="VAT_Monthly")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — MONTHLY VALUE ADDED TAX (VAT) REGISTER").font = title_font
    ws.cell(row=2, column=1, value="Monthly Standard 7.5% Output VAT, Exempt Property Supplies & Remittance Status (2023 - 2025)").font = italic_font
    
    headers = ["Tax Period", "Taxable Commercial Turnover (NGN)", "Exempt Land & Housing Supplies (NGN)", "Total Gross Turnover (NGN)", "Output VAT @ 7.5% (NGN)", "Recoverable Direct Input VAT (NGN)", "Net VAT Payable / (Remitted) (NGN)", "FIRS Remittance Status"]
    style_header(ws, 4, headers)
    
    vat_data = [
        # 2023 Sample Monthly Data
        ("Jan 2023", 1500000.00, 0.00, 1500000.00, 112500.00, 22500.00, 90000.00, "REMITTED"),
        ("Feb 2023", 2000000.00, 0.00, 2000000.00, 150000.00, 30000.00, 120000.00, "REMITTED"),
        ("Mar 2023", 2500000.00, 0.00, 2500000.00, 187500.00, 37500.00, 150000.00, "REMITTED"),
        ("Apr 2023", 1800000.00, 3000000.00, 4800000.00, 135000.00, 27000.00, 108000.00, "REMITTED"),
        ("May 2023", 2200000.00, 0.00, 2200000.00, 165000.00, 33000.00, 132000.00, "REMITTED"),
        ("Jun 2023", 2800000.00, 0.00, 2800000.00, 210000.00, 42000.00, 168000.00, "REMITTED"),
        ("Jul 2023", 2100000.00, 3500000.00, 5600000.00, 157500.00, 31500.00, 126000.00, "REMITTED"),
        ("Aug 2023", 2400000.00, 0.00, 2400000.00, 180000.00, 36000.00, 144000.00, "REMITTED"),
        ("Sep 2023", 2600000.00, 0.00, 2600000.00, 195000.00, 39000.00, 156000.00, "REMITTED"),
        ("Oct 2023", 2000000.00, 3000000.00, 500000.00, 150000.00, 30000.00, 120000.00, "REMITTED"),
        ("Nov 2023", 2100000.00, 0.00, 2100000.00, 157500.00, 31500.00, 126000.00, "REMITTED"),
        ("Dec 2023", 2000000.00, 3000000.00, 5000000.00, 150000.00, 30000.00, 120000.00, "REMITTED"),
        ("TOTAL FY 2023 VAT", "=SUM(B5:B16)", "=SUM(C5:C16)", "=SUM(D5:D16)", "=SUM(E5:E16)", "=SUM(F5:F16)", "=SUM(G5:G16)", "100% RECONCILED"),
    ]
    
    for idx, (mo, tt, et, gt, out_v, in_v, net_v, stat) in enumerate(vat_data, 5):
        is_tot = "TOTAL" in mo
        ws.cell(row=idx, column=1, value=mo).font = bold_font if is_tot else regular_font
        for c_idx, val in enumerate([tt, et, gt, out_v, in_v, net_v], 2):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
        c_st = ws.cell(row=idx, column=8, value=stat)
        c_st.font = success_font if "REMITTED" in stat or "RECONCILED" in stat else regular_font
        c_st.alignment = align_center
        for c in range(1, 9):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

def build_sheet_24_wht_receivable_payable(wb):
    ws = wb.create_sheet(title="WHT_Receivable_Payable")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — WITHHOLDING TAX (WHT) SCHEDULE").font = title_font
    ws.cell(row=2, column=1, value="Client Credit Notes Receivable, Vendor Deductions Payable & CIT Offsets (2023 - 2025)").font = italic_font
    
    headers = ["Contract / Counterparty", "Transaction Description", "Gross Contract Amount (NGN)", "WHT Rate (%)", "WHT Deducted (NGN)", "Credit Note Ref / Status", "Utilization / Offset Against CIT (NGN)", "Balance Carried Forward (NGN)"]
    style_header(ws, 4, headers)
    
    wht_clients = [
        ("Apex Properties Ltd", "Lekki Phase 1 Civil Construction Progress Billings", 22000000.00, 0.05, 1100000.00, "FIRS-CN-2023-0881", 0.00, 1100000.00),
        ("Horizon Holdings Ltd", "Ikeja Commercial Engineering Consulting", 4000000.00, 0.05, 200000.00, "FIRS-CN-2023-1042", 0.00, 200000.00),
        ("TOTAL WHT RECEIVABLE 2023", "Carried Forward as CIT Tax Credit", "=SUM(C5:C6)", None, "=SUM(E5:E6)", "TIED TO SFP NOTE 10", "=SUM(G5:G6)", "=SUM(H5:H6)"),
    ]
    
    for idx, (cp, desc, gross, rate, ded, ref, ut, bal) in enumerate(wht_clients, 5):
        is_tot = "TOTAL" in cp
        ws.cell(row=idx, column=1, value=cp).font = bold_font if is_tot else regular_font
        ws.cell(row=idx, column=2, value=desc).font = regular_font
        
        c_g = ws.cell(row=idx, column=3, value=gross)
        c_g.number_format = fmt_currency
        c_g.alignment = align_right
        c_g.font = bold_font if is_tot else regular_font
        
        if rate:
            c_r = ws.cell(row=idx, column=4, value=rate)
            c_r.number_format = fmt_percent
            c_r.alignment = align_center
        else:
            ws.cell(row=idx, column=4, value="").alignment = align_center
            
        for c_idx, val in enumerate([ded, ref, ut, bal], 5):
            c = ws.cell(row=idx, column=c_idx, value=val)
            if isinstance(val, (int, float)):
                c.number_format = fmt_currency
                c.alignment = align_right
            elif isinstance(val, str) and val.startswith("="):
                c.number_format = fmt_currency
                c.alignment = align_right
            else:
                c.alignment = align_center
            c.font = bold_font if is_tot else regular_font
            
        for c in range(1, 9):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

def build_sheet_25_deferred_tax_ias12(wb):
    ws = wb.create_sheet(title="Deferred_Tax_IAS12")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — DEFERRED TAXATION SCHEDULE (IAS 12)").font = title_font
    ws.cell(row=2, column=1, value="Temporary Differences, Tax Base vs Carrying Amount & Profit or Loss Recognition (2022 - 2025)").font = italic_font
    
    headers = ["Temporary Difference Component", "Carrying Amount (NGN)", "Tax Base (NGN)", "Taxable / (Deductible) Difference (NGN)", "Applicable Tax Rate (%)", "Deferred Tax Asset / (Liability) (NGN)", "Movement in P&L (Credit) / Charge (NGN)"]
    style_header(ws, 4, headers)
    
    dt_rows = [
        ("FY 2023 — Property, Plant & Equipment", 1424540.00, 1289540.00, 135000.00, 0.20, -27000.00, 27000.00),
        ("FY 2023 — IFRS 9 ECL Impairment Allowance", -185000.00, 0.00, -185000.00, 0.20, 37000.00, -37000.00),
        ("FY 2023 — Accrued Expenses & Prepayments", -110000.00, 0.00, -110000.00, 0.20, 22000.00, -22000.00),
        ("FY 2023 — NET DEFERRED TAX ASSET / (LIABILITY)", "=SUM(B5:B7)", "=SUM(C5:C7)", "=SUM(D5:D7)", None, 32000.00, -32000.00),
        
        ("FY 2024 — Property, Plant & Equipment", 3409540.00, 3009540.00, 400000.00, 0.20, -80000.00, 53000.00),
        ("FY 2024 — IFRS 9 ECL Impairment Allowance", -445000.00, 0.00, -445000.00, 0.20, 89000.00, -52000.00),
        ("FY 2024 — Accrued Expenses & Prepayments", -125000.00, 0.00, -125000.00, 0.20, -25000.00, 47000.00),
        ("FY 2024 — NET DEFERRED TAX (LIABILITY)", "=SUM(B9:B11)", "=SUM(C9:C11)", "=SUM(D9:D11)", None, -16000.00, 48000.00),
        
        ("FY 2025 — Property, Plant & Equipment", 5724540.00, 4859540.00, 865000.00, 0.30, -259500.00, 179500.00),
        ("FY 2025 — IFRS 9 ECL Impairment Allowance", -840000.00, 0.00, -840000.00, 0.30, 252000.00, -163000.00),
        ("FY 2025 — Accrued Expenses & Prepayments", -303333.33, 0.00, -303333.33, 0.30, -91000.00, 66000.00),
        ("FY 2025 — NET DEFERRED TAX (LIABILITY)", "=SUM(B13:B15)", "=SUM(C13:C15)", "=SUM(D13:D15)", None, -98500.00, 82500.00),
    ]
    
    for idx, (comp, ca, tb, diff, rate, dta_dtl, pnl) in enumerate(dt_rows, 5):
        is_tot = "NET DEFERRED" in comp
        ws.cell(row=idx, column=1, value=comp).font = bold_font if is_tot else regular_font
        for c_idx, val in enumerate([ca, tb, diff], 2):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
        if rate:
            c_r = ws.cell(row=idx, column=5, value=rate)
            c_r.number_format = fmt_percent
            c_r.alignment = align_center
        else:
            ws.cell(row=idx, column=5, value="").alignment = align_center
        for c_idx, val in enumerate([dta_dtl, pnl], 6):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
        for c in range(1, 8):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

def build_sheet_26_tax_payable_rollforward(wb):
    ws = wb.create_sheet(title="Tax_Payable_Rollforward")
    ws.freeze_panes = "A5"
    
    ws.cell(row=1, column=1, value="BAAY PROJECTS LIMITED — TAX LIABILITIES ROLL-FORWARD SCHEDULE").font = title_font
    ws.cell(row=2, column=1, value="Roll-Forward of All Tax Heads (CIT, TET, VAT, WHT, PTF, NASENI) to Reporting Date & Settlement").font = italic_font
    
    headers = ["Tax Head / Levy", "Opening Liability 1 Jan (NGN)", "Current Tax Provision (NGN)", "Prior Year Adjustments (NGN)", "Tax Remittances / Payments (NGN)", "WHT Credits Utilized (NGN)", "Closing Liability 31 Dec (NGN)", "Settlement & Assessment Status"]
    style_header(ws, 4, headers)
    
    roll_rows = [
        # 2023 Roll-forward
        ("FY 2023 — Companies Income Tax (CIT)", 0.00, 1492000.00, 0.00, 0.00, 0.00, 1492000.00, "Assessed / Payable YOA 2024"),
        ("FY 2023 — Tertiary Education Tax (TET)", 0.00, 239400.00, 0.00, 0.00, 0.00, 239400.00, "Assessed / Payable YOA 2024"),
        ("FY 2023 — Police Trust Fund (PTF)", 90.00, 365.00, 0.00, 0.00, 0.00, 455.00, "Includes ₦90 opening arrears"),
        ("FY 2023 — TOTAL CURRENT TAX LIABILITIES", 90.00, 1731765.00, 0.00, 0.00, 0.00, 1731855.00, "TIED TO SFP NOTE 16"),
        
        # 2024 Roll-forward
        ("FY 2024 — Companies Income Tax (CIT)", 1492000.00, 2528000.00, 0.00, -1492000.00, -400000.00, 2128000.00, "2023 CIT Paid; WHT Credit Applied"),
        ("FY 2024 — Tertiary Education Tax (TET)", 239400.00, 408600.00, 0.00, -239400.00, 0.00, 408600.00, "2023 TET Remitted"),
        ("FY 2024 — Police Trust Fund (PTF)", 455.00, 625.00, 0.00, -455.00, 0.00, 625.00, "Opening PTF Arrears Settled"),
        ("FY 2024 — TOTAL CURRENT TAX LIABILITIES", 1731855.00, 2937225.00, 0.00, -1731855.00, -400000.00, 2537225.00, "TIED TO SFP NOTE 16"),
        
        # 2025 Roll-forward
        ("FY 2025 — Companies Income Tax (CIT)", 2128000.00, 6825000.00, 0.00, -2128000.00, -2000000.00, 4825000.00, "2024 CIT Paid; ₦2m WHT Utilized"),
        ("FY 2025 — Tertiary Education Tax (TET)", 408600.00, 735000.00, 0.00, -408600.00, 0.00, 735000.00, "2024 TET Remitted"),
        ("FY 2025 — Police Trust Fund (PTF)", 625.00, 1130.00, 0.00, -625.00, 0.00, 1130.00, "2024 PTF Remitted"),
        ("FY 2025 — NASENI Development Levy", 0.00, 56500.00, 0.00, 0.00, 0.00, 56500.00, "0.25% Large Company Levy"),
        ("FY 2025 — TOTAL CURRENT TAX LIABILITIES", 2537225.00, 7617630.00, 0.00, -2537225.00, -2000000.00, 5617630.00, "TIED TO SFP NOTE 16"),
    ]
    
    for idx, (th, op, prov, adj, rem, wht_u, cl, stat) in enumerate(roll_rows, 5):
        is_tot = "TOTAL" in th
        ws.cell(row=idx, column=1, value=th).font = bold_font if is_tot else regular_font
        for c_idx, val in enumerate([op, prov, adj, rem, wht_u, cl], 2):
            c = ws.cell(row=idx, column=c_idx, value=val)
            c.number_format = fmt_currency
            c.alignment = align_right
            c.font = bold_font if is_tot else regular_font
        c_st = ws.cell(row=idx, column=8, value=stat)
        c_st.font = success_font if is_tot else italic_font
        c_st.alignment = align_center if is_tot else align_left
        for c in range(1, 9):
            ws.cell(row=idx, column=c).border = double_bottom if is_tot else thin_border
            if is_tot: ws.cell(row=idx, column=c).fill = total_fill
            elif idx % 2 == 1: ws.cell(row=idx, column=c).fill = zebra_fill
            
    auto_fit(ws)

print("sheet_builders_4.py completed with Sheets 20-26.")
