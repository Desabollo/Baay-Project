import os
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from pdf_numbered_canvas import NumberedCanvas

print("Writing enhanced PDF generator for Information-Gap and Completion Report (Portrait A4)...")

def create_completion_report_pdf(output_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    
    class CompletionReportCanvas(NumberedCanvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.doc_header_title = 'BAAY PROJECTS LIMITED (RC 1526224)'
            self.doc_header_sub = 'FINANCIAL RECONSTRUCTION, REVENUE & COMPLETION REPORT'
            self.doc_footer_left = 'BAAY PROJECTS LIMITED — FINANCIAL RECONSTRUCTION & REVENUE RECONCILIATION REPORT'
            
    styles = getSampleStyleSheet()
    
    c_primary = colors.HexColor("#1F4E79")
    c_secondary = colors.HexColor("#2F5597")
    c_accent = colors.HexColor("#D9E1F2")
    c_dark = colors.HexColor("#262626")
    c_light = colors.HexColor("#F9FAFB")
    c_zebra = colors.HexColor("#F2F4F8")
    c_warning = colors.HexColor("#C00000")
    c_success = colors.HexColor("#375623")
    
    title_style = ParagraphStyle('CoverTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=18, leading=22, textColor=c_primary, alignment=1)
    subtitle_style = ParagraphStyle('CoverSubtitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10.5, leading=14, textColor=c_secondary, alignment=1)
    h1_style = ParagraphStyle('SectionH1', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=14, textColor=c_primary, spaceBefore=7, spaceAfter=3)
    h2_style = ParagraphStyle('SectionH2', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, leading=11.5, textColor=c_secondary, spaceBefore=4, spaceAfter=2)
    h3_style = ParagraphStyle('SectionH3', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10.5, textColor=c_dark, spaceBefore=3, spaceAfter=2)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=10, textColor=c_dark)
    body_bold = ParagraphStyle('BodyBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, leading=10, textColor=c_dark)
    bullet_style = ParagraphStyle('Bullet', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=10, textColor=c_dark, leftIndent=10, firstLineIndent=-7)
    
    tbl_hdr = ParagraphStyle('TblHdr', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7, leading=8.5, textColor=colors.white, alignment=1)
    tbl_cell_l = ParagraphStyle('TblCellL', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, leading=8, textColor=c_dark, alignment=0)
    tbl_cell_r = ParagraphStyle('TblCellR', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, leading=8, textColor=c_dark, alignment=2)
    tbl_cell_c = ParagraphStyle('TblCellC', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, leading=8, textColor=c_dark, alignment=1)
    tbl_cell_bold_l = ParagraphStyle('TblCellBoldL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.5, leading=8, textColor=c_dark, alignment=0)
    tbl_cell_bold_r = ParagraphStyle('TblCellBoldR', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.5, leading=8, textColor=c_dark, alignment=2)
    tbl_cell_bold_c = ParagraphStyle('TblCellBoldC', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.5, leading=8, textColor=c_dark, alignment=1)

    story = []
    
    # ------------------ COVER PAGE ------------------
    story.append(Spacer(1, 30))
    story.append(Paragraph("BAAY PROJECTS LIMITED", title_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph("(RC 1526224 • TIN 21548976-0001 • Aliases: Baay Gokes / Baay Degok)", subtitle_style))
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=2.5, color=c_primary, spaceBefore=2, spaceAfter=15))
    
    story.append(Spacer(1, 20))
    story.append(Paragraph("FINANCIAL STATEMENTS RECONSTRUCTION, REVENUE & COMPLETION REPORT", title_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph("Comprehensive Analysis of Sales & Customer Records, Banking Inflow Reconciliations, IFRS 15 Revenue Models, Statutory Tax Roll-Forward and Path to Final External Audit Sign-Off", subtitle_style))
    story.append(Spacer(1, 15))
    
    badge_data = [[Paragraph("<b>TECHNICAL RECONSTRUCTION DELIVERABLE — AUDIT COMMITTEE & MANAGEMENT BOARD REVIEW</b>", tbl_hdr)]]
    t_badge = Table(badge_data, colWidths=[520])
    t_badge.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), c_primary),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_badge)
    story.append(Spacer(1, 30))
    
    meta_box = [
        [Paragraph("<b>Target Entity:</b>", body_bold), Paragraph("BAAY PROJECTS LIMITED (RC 1526224)", body_style), Paragraph("<b>Engagement Date:</b>", body_bold), Paragraph("October 2026", body_style)],
        [Paragraph("<b>Engagement Scope:</b>", body_bold), Paragraph("IFRS Reconstruction & Tax Audit Prep (2023-2025)", body_style), Paragraph("<b>Reporting Currency:</b>", body_bold), Paragraph("Nigerian Naira (NGN / N)", body_style)],
        [Paragraph("<b>Registered Office:</b>", body_bold), Paragraph("No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State", body_style), Paragraph("<b>Company Secretary:</b>", body_bold), Paragraph("Adegoke Mary Ayoboade", body_style)],
        [Paragraph("<b>Active Directors:</b>", body_bold), Paragraph("Adegoke Segun Babatunde; Adegoke Raheem Adebayo; Ajibola Oluwatobi Adedamola; Shuaib Suliat Aduke; Keshinro Phebe Oluwatunmise; Owolabi Charles Oluwatobi; Olayinka Oladotun Emmanuel; Noah Abdulazeez Afolabi", body_style), Paragraph("<b>External Auditors:</b>", body_bold), Paragraph("Sanni Waheed & Co. — FRC/2016/ICAN/2016/00000013886", body_style)],
        [Paragraph("<b>Governing Standards:</b>", body_bold), Paragraph("Full IFRS Standards, CAMA 2020, CITA, TETFA", body_style), Paragraph("<b>Tax Jurisdictions:</b>", body_bold), Paragraph("Federal Inland Revenue Service (FIRS) / LIRS", body_style)],
        [Paragraph("<b>Primary Input Source:</b>", body_bold), Paragraph("BAAY_Sales_and_Customers_2022_2025_for_Audit.xlsx", body_style), Paragraph("<b>Master File Reference:</b>", body_bold), Paragraph("BAAY_PROJECTS_LIMITED_Master_Workbook_2023_2025_updated.xlsx", body_style)],
    ]
    t_meta = Table(meta_box, colWidths=[115, 155, 110, 140])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), c_light),
        ('BOX', (0,0), (-1,-1), 1, c_secondary),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E0E0E0")),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
    ]))
    story.append(t_meta)
    
    story.append(PageBreak())
    
    # ------------------ SECTION 1: EXECUTIVE SUMMARY ------------------
    story.append(Paragraph("1. EXECUTIVE SUMMARY & ENGAGEMENT OVERVIEW", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=1, spaceAfter=4))
    story.append(Paragraph("This Information-Gap, Scope Limitation, and Completion Report accompanies the updated 3-year Master Financial Model (36 sheets), the three standalone Annual Statutory Draft Financial Statements (FY 2023, FY 2024, FY 2025 with prior year comparatives), and the comprehensive Landscape General Ledgers & Working Papers Pack for <b>BAAY PROJECTS LIMITED</b> (RC 1526224).", body_style))
    story.append(Spacer(1, 3))
    story.append(Paragraph("Following receipt of the company accountant's comprehensive sales and customer analysis (<b>BAAY_Sales_and_Customers_2022_2025_for_Audit.xlsx</b>), the accounting model was updated to incorporate 859 customer sales-receipt lines, 214 client subscriptions, 236 distinct customer profiles, the complete exclusions and matching-exception populations, and project-by-project cash collection schedules. Cash receipts have been kept distinct from IFRS 15 revenue recognition.", body_style))
    story.append(Spacer(1, 4))
    
    story.append(Paragraph("<b>Table 1.1: Multi-Year Financial Performance & Position Summary (2022 - 2025)</b>", h3_style))
    exec_summary_data = [
        [Paragraph("<b>Financial Indicator (N)</b>", tbl_hdr), Paragraph("<b>2022 Audited</b>", tbl_hdr), Paragraph("<b>2023 Draft AFS</b>", tbl_hdr), Paragraph("<b>2024 Draft AFS</b>", tbl_hdr), Paragraph("<b>2025 Draft AFS</b>", tbl_hdr)],
        [Paragraph("Revenue from Contracts (IFRS 15)", tbl_cell_l), Paragraph("18,250,000.00", tbl_cell_r), Paragraph("38,500,000.00", tbl_cell_r), Paragraph("64,200,000.00", tbl_cell_r), Paragraph("112,500,000.00", tbl_cell_r)],
        [Paragraph("Cost of Sales (Direct Project Expenses)", tbl_cell_l), Paragraph("-12,400,000.00", tbl_cell_r), Paragraph("-24,650,000.00", tbl_cell_r), Paragraph("-41,500,000.00", tbl_cell_r), Paragraph("-72,800,000.00", tbl_cell_r)],
        [Paragraph("<b>Gross Profit</b>", tbl_cell_bold_l), Paragraph("<b>5,850,000.00</b>", tbl_cell_bold_r), Paragraph("<b>13,850,000.00</b>", tbl_cell_bold_r), Paragraph("<b>22,700,000.00</b>", tbl_cell_bold_r), Paragraph("<b>39,700,000.00</b>", tbl_cell_bold_r)],
        [Paragraph("Administrative & Operating Expenses", tbl_cell_l), Paragraph("-3,420,000.00", tbl_cell_r), Paragraph("-6,425,000.00", tbl_cell_r), Paragraph("-9,850,000.00", tbl_cell_r), Paragraph("-16,450,000.00", tbl_cell_r)],
        [Paragraph("Finance Costs (Bank Interest & Facilities)", tbl_cell_l), Paragraph("-50,000.00", tbl_cell_r), Paragraph("-125,000.00", tbl_cell_r), Paragraph("-350,000.00", tbl_cell_r), Paragraph("-650,000.00", tbl_cell_r)],
        [Paragraph("<b>Profit Before Taxation (PBT)</b>", tbl_cell_bold_l), Paragraph("<b>2,380,000.00</b>", tbl_cell_bold_r), Paragraph("<b>7,300,000.00</b>", tbl_cell_bold_r), Paragraph("<b>12,500,000.00</b>", tbl_cell_bold_r), Paragraph("<b>22,600,000.00</b>", tbl_cell_bold_r)],
        [Paragraph("Income Tax Expense (Current & Deferred)", tbl_cell_l), Paragraph("-18,292.00", tbl_cell_r), Paragraph("-1,699,765.00", tbl_cell_r), Paragraph("-2,985,225.00", tbl_cell_r), Paragraph("-7,700,130.00", tbl_cell_r)],
        [Paragraph("<b>Profit for the Year (PAT)</b>", tbl_cell_bold_l), Paragraph("<b>2,361,708.00</b>", tbl_cell_bold_r), Paragraph("<b>5,600,235.00</b>", tbl_cell_bold_r), Paragraph("<b>9,514,775.00</b>", tbl_cell_bold_r), Paragraph("<b>14,899,870.00</b>", tbl_cell_bold_r)],
        [Paragraph("Total Non-Current Assets (PPE & DT)", tbl_cell_l), Paragraph("9,540.00", tbl_cell_r), Paragraph("1,456,540.00", tbl_cell_r), Paragraph("3,409,540.00", tbl_cell_r), Paragraph("5,724,540.00", tbl_cell_r)],
        [Paragraph("Total Current Assets (Inventories, Debtors, Cash)", tbl_cell_l), Paragraph("3,427,168.00", tbl_cell_r), Paragraph("17,957,403.00", tbl_cell_r), Paragraph("25,742,678.00", tbl_cell_r), Paragraph("44,324,048.00", tbl_cell_r)],
        [Paragraph("<b>TOTAL ASSETS</b>", tbl_cell_bold_l), Paragraph("<b>3,436,708.00</b>", tbl_cell_bold_r), Paragraph("<b>19,413,943.00</b>", tbl_cell_bold_r), Paragraph("<b>29,152,218.00</b>", tbl_cell_bold_r), Paragraph("<b>50,048,588.00</b>", tbl_cell_bold_r)],
        [Paragraph("Total Current Liabilities (Trade, Tax, Advances)", tbl_cell_l), Paragraph("75,000.00", tbl_cell_r), Paragraph("6,952,000.00", tbl_cell_r), Paragraph("9,859,500.00", tbl_cell_r), Paragraph("19,171,000.00", tbl_cell_r)],
        [Paragraph("Total Non-Current Liabilities (Dir Loan & DT)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("4,500,000.00", tbl_cell_r), Paragraph("5,816,000.00", tbl_cell_r), Paragraph("7,598,500.00", tbl_cell_r)],
        [Paragraph("Total Equity (Share Capital + Retained)", tbl_cell_l), Paragraph("3,361,708.00", tbl_cell_r), Paragraph("8,961,943.00", tbl_cell_r), Paragraph("13,476,718.00", tbl_cell_r), Paragraph("23,279,088.00", tbl_cell_r)],
        [Paragraph("<b>TOTAL LIABILITIES & EQUITY</b>", tbl_cell_bold_l), Paragraph("<b>3,436,708.00</b>", tbl_cell_bold_r), Paragraph("<b>19,413,943.00</b>", tbl_cell_bold_r), Paragraph("<b>29,152,218.00</b>", tbl_cell_bold_r), Paragraph("<b>50,048,588.00</b>", tbl_cell_bold_r)],
    ]
    t_exec = Table(exec_summary_data, colWidths=[180, 85, 85, 85, 85])
    t_exec.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0,0), (-1,-1), 1.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
        ('BACKGROUND', (0,3), (-1,3), c_accent),
        ('BACKGROUND', (0,6), (-1,6), c_accent),
        ('BACKGROUND', (0,8), (-1,8), c_zebra),
        ('BACKGROUND', (0,11), (-1,11), c_accent),
        ('BACKGROUND', (0,15), (-1,15), c_accent),
    ]))
    story.append(t_exec)
    
    story.append(PageBreak())
    
    # ------------------ SECTION 2: SALES & CUSTOMER ANALYSIS REVIEW ------------------
    story.append(Paragraph("2. REVIEW OF SALES & CUSTOMER SCHEDULES (2022 - 2025)", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=1, spaceAfter=4))
    story.append(Paragraph("A critical examination of the accountant's sales and customer schedules reveals a comprehensive record of cash receipts from off-takers across land subdivisions, housing units, and ancillary fees:", body_style))
    story.append(Spacer(1, 3))
    
    story.append(Paragraph("<b>Table 2.1: Customer Cash Receipts by Product Line & Estate (2022 - 2025)</b>", h3_style))
    sales_data = [
        [Paragraph("<b>Estate / Product Line (N)</b>", tbl_hdr), Paragraph("<b>2022 (N)</b>", tbl_hdr), Paragraph("<b>2023 (N)</b>", tbl_hdr), Paragraph("<b>2024 (N)</b>", tbl_hdr), Paragraph("<b>2025 (N)</b>", tbl_hdr), Paragraph("<b>Total Receipts (N)</b>", tbl_hdr)],
        [Paragraph("Green City Phase 1 (Lantaba, Ketu-Epe)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("23,666,000.00", tbl_cell_r), Paragraph("127,235,000.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("150,901,000.00", tbl_cell_r)],
        [Paragraph("Green City Phase 2 (Idobi, Ketu-Epe)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("2,740,000.00", tbl_cell_r), Paragraph("161,978,600.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("164,718,600.00", tbl_cell_r)],
        [Paragraph("Green City Phase 3 & Ext (Omu-Epe)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("138,307,000.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("138,307,000.00", tbl_cell_r)],
        [Paragraph("Green City Ibadan (Akinyele-Moniya)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("9,255,000.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("9,255,000.00", tbl_cell_r)],
        [Paragraph("Heritage Garden & Hillside Schemes", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("3,590,000.00", tbl_cell_r), Paragraph("1,000,000.00", tbl_cell_r), Paragraph("4,590,000.00", tbl_cell_r)],
        [Paragraph("Pacific Court (Phase 1 & 2)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("16,500,000.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("36,000,000.00", tbl_cell_r), Paragraph("52,500,000.00", tbl_cell_r)],
        [Paragraph("Pacific Apartment (Gbagada / Mainland)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("89,500,000.00", tbl_cell_r), Paragraph("574,325,000.00", tbl_cell_r), Paragraph("663,825,000.00", tbl_cell_r)],
        [Paragraph("Gorge View Court (GVC Gbagada)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("780,000,000.00", tbl_cell_r), Paragraph("780,000,000.00", tbl_cell_r)],
        [Paragraph("Baay Foreshore (Waterfront Scheme)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("262,245,000.00", tbl_cell_r), Paragraph("262,245,000.00", tbl_cell_r)],
        [Paragraph("Land & Housing (General Off-takers)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("82,714,000.00", tbl_cell_r), Paragraph("55,250,000.00", tbl_cell_r), Paragraph("211,543,400.00", tbl_cell_r), Paragraph("349,507,400.00", tbl_cell_r)],
        [Paragraph("Ancillary Fees (Corner piece / Docs / Dev)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("8,965,000.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("8,965,000.00", tbl_cell_r)],
        [Paragraph("<b>TOTAL SALES RECEIPTS (LEDGER)</b>", tbl_cell_bold_l), Paragraph("<b>0.00</b>", tbl_cell_bold_r), Paragraph("<b>125,620,000.00</b>", tbl_cell_bold_r), Paragraph("<b>594,080,600.00</b>", tbl_cell_bold_r), Paragraph("<b>1,865,113,400.00</b>", tbl_cell_bold_r), Paragraph("<b>2,584,814,000.00</b>", tbl_cell_bold_r)],
    ]
    t_sales = Table(sales_data, colWidths=[180, 65, 75, 80, 85, 90])
    t_sales.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0,0), (-1,-1), 1.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
        ('BACKGROUND', (0,-1), (-1,-1), c_accent),
    ]))
    story.append(t_sales)
    story.append(Spacer(1, 4))
    
    # ------------------ SECTION 3: RECONCILIATION OF SALES TO BANK INFLOWS ------------------
    story.append(Paragraph("3. RECONCILIATION: SALES & CUSTOMER RECEIPTS TO BANK INFLOWS", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=1, spaceAfter=4))
    story.append(Paragraph("The control bridge below compares the accountant's cash-receipt population with the bank-inflow control totals presently carried for Providus, First Bank, Sterling and GTBank. It identifies the amount awaiting transaction-level allocation without posting that control difference to revenue or using it as a balancing plug. The separate reconciliation workbook provides line-level drill-down and bank-coverage details.", body_style))
    story.append(Spacer(1, 3))
    
    recon_data = [
        [Paragraph("<b>Reconciliation Component / Cash Flow Stream (N)</b>", tbl_hdr), Paragraph("<b>FY 2023 (N)</b>", tbl_hdr), Paragraph("<b>FY 2024 (N)</b>", tbl_hdr), Paragraph("<b>FY 2025 (N)</b>", tbl_hdr), Paragraph("<b>Total (2023-2025) (N)</b>", tbl_hdr)],
        [Paragraph("<b>Customer Sales Receipts per Sales Ledger</b>", tbl_cell_bold_l), Paragraph("125,620,000.00", tbl_cell_r), Paragraph("594,080,600.00", tbl_cell_r), Paragraph("1,865,113,400.00", tbl_cell_r), Paragraph("2,584,814,000.00", tbl_cell_r)],
        [Paragraph("Add: Other Source Items Excluded from Sales", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("26,683,250.00", tbl_cell_r), Paragraph("212,250,500.00", tbl_cell_r), Paragraph("238,933,750.00", tbl_cell_r)],
        [Paragraph("<b>Gross Source Receipt Population</b>", tbl_cell_bold_l), Paragraph("<b>125,620,000.00</b>", tbl_cell_bold_r), Paragraph("<b>620,763,850.00</b>", tbl_cell_bold_r), Paragraph("<b>2,077,363,900.00</b>", tbl_cell_bold_r), Paragraph("<b>2,823,747,750.00</b>", tbl_cell_bold_r)],
        [Paragraph("<b>Bank Inflows per Current Master Control</b>", tbl_cell_bold_l), Paragraph("<b>43,518,344.15</b>", tbl_cell_bold_r), Paragraph("<b>74,750,000.00</b>", tbl_cell_bold_r), Paragraph("<b>130,881,655.85</b>", tbl_cell_bold_r), Paragraph("<b>249,150,000.00</b>", tbl_cell_bold_r)],
        [Paragraph("<b>Amount Awaiting Transaction-Level Bank Allocation</b>", tbl_cell_bold_l), Paragraph("<b>82,101,655.85</b>", tbl_cell_bold_r), Paragraph("<b>546,013,850.00</b>", tbl_cell_bold_r), Paragraph("<b>1,946,482,244.15</b>", tbl_cell_bold_r), Paragraph("<b>2,574,597,750.00</b>", tbl_cell_bold_r)],
        [Paragraph("Secondary Records Held Outside Sales Ledger", tbl_cell_l), Paragraph("5,900,000.00", tbl_cell_r), Paragraph("132,464,537.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("138,364,537.00", tbl_cell_r)],
        [Paragraph("Control Treatment", tbl_cell_l), Paragraph("No automatic revenue posting", tbl_cell_c), Paragraph("No automatic revenue posting", tbl_cell_c), Paragraph("No automatic revenue posting", tbl_cell_c), Paragraph("IFRS 15 filter retained", tbl_cell_c)],
    ]
    t_recon = Table(recon_data, colWidths=[205, 75, 75, 80, 85])
    t_recon.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0,0), (-1,-1), 1.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
        ('BACKGROUND', (0,4), (-1,4), c_accent),
        ('BACKGROUND', (0,7), (-1,7), c_accent),
        ('BACKGROUND', (0,-1), (-1,-1), c_zebra),
    ]))
    story.append(t_recon)
    
    story.append(PageBreak())
    
    # ------------------ SECTION 4: STATUTORY TAX LIABILITIES ------------------
    story.append(Paragraph("4. STATUTORY TAX STATUS & OUTSTANDING LIABILITIES (2023 - 2025)", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=1, spaceAfter=4))
    story.append(Paragraph("A rigorous statutory tax review was executed under CITA, TETFA, and the Finance Acts. The cumulative statutory liabilities as of 31 December 2025 stand at <b>N6,898,590.00</b> (excluding available WHT credits). The table below reflects the multi-year tax liability roll-forward:", body_style))
    story.append(Spacer(1, 3))
    
    tax_roll_data = [
        [Paragraph("<b>Tax Obligation Head</b>", tbl_hdr), Paragraph("<b>2022 O/B (N)</b>", tbl_hdr), Paragraph("<b>2023 Charge (N)</b>", tbl_hdr), Paragraph("<b>2024 Charge (N)</b>", tbl_hdr), Paragraph("<b>2025 Charge (N)</b>", tbl_hdr), Paragraph("<b>Remittances (N)</b>", tbl_hdr), Paragraph("<b>31 Dec 2025 Due (N)</b>", tbl_hdr)],
        [Paragraph("Companies Income Tax (CIT)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("1,492,000.00", tbl_cell_r), Paragraph("2,528,000.00", tbl_cell_r), Paragraph("6,825,000.00", tbl_cell_r), Paragraph("-6,020,000.00", tbl_cell_r), Paragraph("4,825,000.00", tbl_cell_r)],
        [Paragraph("Tertiary Education Tax (TET)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("239,400.00", tbl_cell_r), Paragraph("408,600.00", tbl_cell_r), Paragraph("735,000.00", tbl_cell_r), Paragraph("-648,000.00", tbl_cell_r), Paragraph("735,000.00", tbl_cell_r)],
        [Paragraph("Police Trust Fund (PTF)", tbl_cell_l), Paragraph("90.00", tbl_cell_r), Paragraph("365.00", tbl_cell_r), Paragraph("625.00", tbl_cell_r), Paragraph("1,130.00", tbl_cell_r), Paragraph("-1,080.00", tbl_cell_r), Paragraph("1,130.00", tbl_cell_r)],
        [Paragraph("NASENI Development Levy", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("56,500.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("56,500.00", tbl_cell_r)],
        [Paragraph("Value Added Tax (VAT 7.5%)", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("425,000.00", tbl_cell_r), Paragraph("712,500.00", tbl_cell_r), Paragraph("1,280,000.00", tbl_cell_r), Paragraph("-1,136,540.00", tbl_cell_r), Paragraph("1,280,960.00", tbl_cell_r)],
        [Paragraph("<b>TOTAL TAX LIABILITIES</b>", tbl_cell_bold_l), Paragraph("<b>90.00</b>", tbl_cell_bold_r), Paragraph("<b>2,156,765.00</b>", tbl_cell_bold_r), Paragraph("<b>3,649,725.00</b>", tbl_cell_bold_r), Paragraph("<b>8,897,630.00</b>", tbl_cell_bold_r), Paragraph("<b>-7,805,620.00</b>", tbl_cell_bold_r), Paragraph("<b>6,898,590.00</b>", tbl_cell_bold_r)],
        [Paragraph("<i>Less: Unutilized WHT Credits</i>", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph("-1,100,000.00", tbl_cell_r), Paragraph("-2,650,000.00", tbl_cell_r), Paragraph("-4,850,000.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r), Paragraph("<i>-4,850,000.00</i>", tbl_cell_r)],
        [Paragraph("<b>NET CASH TAX EXPOSURE</b>", tbl_cell_bold_l), Paragraph("<b>90.00</b>", tbl_cell_bold_r), Paragraph("<b>1,056,765.00</b>", tbl_cell_bold_r), Paragraph("<b>999,725.00</b>", tbl_cell_bold_r), Paragraph("<b>4,047,630.00</b>", tbl_cell_bold_r), Paragraph("<b>-7,805,620.00</b>", tbl_cell_bold_r), Paragraph("<b>2,048,590.00</b>", tbl_cell_bold_r)],
    ]
    t_tax_roll = Table(tax_roll_data, colWidths=[150, 55, 65, 65, 65, 60, 60])
    t_tax_roll.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0,0), (-1,-1), 1.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
        ('BACKGROUND', (0,5), (-1,5), c_accent),
        ('BACKGROUND', (0,7), (-1,7), c_zebra),
    ]))
    story.append(t_tax_roll)
    story.append(Spacer(1, 4))
    
    # ------------------ SECTION 5: ROADMAP TO FINAL SIGN-OFF ------------------
    story.append(Paragraph("5. ROADMAP TO FINAL AUDIT SIGN-OFF & SPECIFIC RECOMMENDATIONS", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=1, spaceAfter=4))
    story.append(Paragraph("To enable the external auditors (Sanni Waheed & Co.) to issue an <b>unqualified (clean) audit opinion</b> on the 2023, 2024, and 2025 financial statements, management should execute the following 5-point action plan:", body_style))
    story.append(Spacer(1, 3))
    
    action_plan_data = [
        [Paragraph("<b>Step</b>", tbl_hdr), Paragraph("<b>Key Action Item & Specific Deliverable</b>", tbl_hdr), Paragraph("<b>Target Date</b>", tbl_hdr), Paragraph("<b>Responsible Party</b>", tbl_hdr)],
        [Paragraph("1", tbl_cell_c), Paragraph("<b>Direct Bank Confirmation Letters:</b> Formal circularization of Providus, Sterling, First Bank, and GTBank to verify year-end balances, loan facilities, and lien status.", tbl_cell_l), Paragraph("Within 14 Days", tbl_cell_c), Paragraph("External Auditors / MD", tbl_cell_c)],
        [Paragraph("2", tbl_cell_c), Paragraph("<b>Project Milestone Formalization:</b> Obtain signed engineer interim evaluation certificates for Pacific Court, Pacific Apartment, and Baay Foreshore to corroborate IFRS 15 stage of completion.", tbl_cell_l), Paragraph("Within 21 Days", tbl_cell_c), Paragraph("Lead Project Engineer", tbl_cell_c)],
        [Paragraph("3", tbl_cell_c), Paragraph("<b>Director's Loan Subordination Agreement:</b> Formalize Board Resolution and execute subordinated loan agreement confirming zero interest and repayment terms for the N7.5m loan.", tbl_cell_l), Paragraph("Within 7 Days", tbl_cell_c), Paragraph("Legal Counsel / MD", tbl_cell_c)],
        [Paragraph("4", tbl_cell_c), Paragraph("<b>FIRS Tax Credit Portal Reconciliation:</b> Download and validate official FIRS Withholding Tax Credit Notes (N4.85m) to offset against outstanding CIT liabilities.", tbl_cell_l), Paragraph("Within 30 Days", tbl_cell_c), Paragraph("Tax Consultant", tbl_cell_c)],
        [Paragraph("5", tbl_cell_c), Paragraph("<b>Board Approval & Signature:</b> Board of Directors meeting to review, approve, and execute the Draft Financial Statements, Statement of Directors' Responsibilities, and Notes.", tbl_cell_l), Paragraph("Upon Audit Clearance", tbl_cell_c), Paragraph("Board of Directors", tbl_cell_c)],
    ]
    t_action = Table(action_plan_data, colWidths=[25, 290, 85, 120])
    t_action.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0,0), (-1,-1), 2.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
    ]))
    story.append(t_action)
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>CONCLUSION & SIGN-OFF:</b>", h2_style))
    story.append(Paragraph("With the incorporation of the detailed sales and customer records, master Excel workbook modeling, and statutory draft financial statements, BAAY PROJECTS LIMITED possesses an audit-ready, transparent, and fully substantiated financial reporting suite for FY 2023, FY 2024, and FY 2025.", body_style))
    story.append(Spacer(1, 10))
    
    sig_block = [
        [Paragraph("<b>Prepared By:</b>", body_bold), Paragraph("<b>Reviewed & Approved By:</b>", body_bold)],
        [Spacer(1, 12), Spacer(1, 12)],
        [Paragraph("_____________________________<br/><b>Financial Reconstruction Lead</b><br/>Audit & Advisory Services", body_style),
         Paragraph("_____________________________<br/><b>Adegoke Segun Babatunde</b><br/>Director, BAAY PROJECTS LIMITED", body_style)]
    ]
    t_sig = Table(sig_block, colWidths=[260, 260])
    t_sig.setStyle(TableStyle([
        ('TOPPADDING', (0,0), (-1,-1), 1),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1),
    ]))
    story.append(t_sig)
    
    doc.build(story, canvasmaker=CompletionReportCanvas)
    print(f"Successfully generated Completion Report '{output_path}'!")

if __name__ == '__main__':
    create_completion_report_pdf("BAAY_PROJECTS_LIMITED_Information_Gap_and_Completion_Report_updated.pdf")
