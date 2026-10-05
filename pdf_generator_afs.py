import os
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from pdf_numbered_canvas import NumberedCanvas

print("Writing PDF generator for Annual Audited Financial Statements...")

def create_afs_pdf(year, output_path):
    # Year-specific setup
    comp_year = year - 1
    
    # Financial data mappings
    if year == 2023:
        p_rev, c_rev = 38500000.00, 8500000.00
        p_cos, c_cos = -24650000.00, -5200000.00
        p_gp, c_gp = 13850000.00, 3300000.00
        p_opx, c_opx = -6425000.00, -2588292.00
        p_ebit, c_ebit = 7425000.00, 711708.00
        p_fin, c_fin = -125000.00, 0.00
        p_pbt, c_pbt = 7300000.00, 711708.00
        p_tax, c_tax = -1699765.00, 0.00
        p_pat, c_pat = 5600235.00, 711708.00
        
        # SFP
        p_ppe, c_ppe = 1424540.00, 9540.00
        p_dta, c_dta = 32000.00, 0.00
        p_nca, c_nca = 1456540.00, 9540.00
        
        p_inv, c_inv = 8650000.00, 0.00
        p_ca, c_ca = 2400000.00, 0.00
        p_rec, c_rec = 5015000.00, 3374668.00
        p_csh, c_csh = 2842403.00, 52500.00
        p_cur_a, c_cur_a = 18907403.00, 3427168.00
        p_tot_a, c_tot_a = 20363943.00, 3436708.00
        
        p_sc, c_sc = 1000000.00, 1000000.00
        p_re, c_re = 7961943.00, 2361708.00
        p_eq, c_eq = 8961943.00, 3361708.00
        
        p_bor, c_bor = 4500000.00, 0.00
        p_dtl, c_dtl = 0.00, 0.00
        p_ncl, c_ncl = 4500000.00, 0.00
        
        p_pay, c_pay = 2770235.00, 75000.00
        p_cl, c_cl = 2400000.00, 0.00
        p_ctx, c_ctx = 1731765.00, 0.00
        p_cur_l, c_cur_l = 6902000.00, 75000.00
        p_tot_l_eq, c_tot_l_eq = 20363943.00, 3436708.00
        
        cit_rate_str = "20% (Medium Company)"
        tet_rate_str = "3% (Finance Act 2023)"
        
    elif year == 2024:
        p_rev, c_rev = 64200000.00, 38500000.00
        p_cos, c_cos = -41500000.00, -24650000.00
        p_gp, c_gp = 22700000.00, 13850000.00
        p_opx, c_opx = -9850000.00, -6425000.00
        p_ebit, c_ebit = 12850000.00, 7425000.00
        p_fin, c_fin = -350000.00, -125000.00
        p_pbt, c_pbt = 12500000.00, 7300000.00
        p_tax, c_tax = -2985225.00, -1699765.00
        p_pat, c_pat = 9514775.00, 5600235.00
        
        p_ppe, c_ppe = 3409540.00, 1424540.00
        p_dta, c_dta = 0.00, 32000.00
        p_nca, c_nca = 3409540.00, 1456540.00
        
        p_inv, c_inv = 13800000.00, 8650000.00
        p_ca, c_ca = 4150000.00, 2400000.00
        p_rec, c_rec = 8685000.00, 5015000.00
        p_csh, c_csh = 5142678.00, 2842403.00
        p_cur_a, c_cur_a = 31777678.00, 18907403.00
        p_tot_a, c_tot_a = 35187218.00, 20363943.00
        
        p_sc, c_sc = 1000000.00, 1000000.00
        p_re, c_re = 17476718.00, 7961943.00
        p_eq, c_eq = 18476718.00, 8961943.00
        
        p_bor, c_bor = 5800000.00, 4500000.00
        p_dtl, c_dtl = 16000.00, 0.00
        p_ncl, c_ncl = 5816000.00, 450000.00
        
        p_pay, c_pay = 4825500.00, 2770235.00
        p_cl, c_cl = 3600000.00, 2400000.00
        p_ctx, c_ctx = 2469000.00, 1731765.00
        p_cur_l, c_cur_l = 10894500.00, 6902000.00
        p_tot_l_eq, c_tot_l_eq = 35187218.00, 20363943.00
        
        cit_rate_str = "20% (Medium Company)"
        tet_rate_str = "3% (Finance Act 2023)"
        
    else: # 2025
        p_rev, c_rev = 112500000.00, 64200000.00
        p_cos, c_cos = -72800000.00, -41500000.00
        p_gp, c_gp = 39700000.00, 22700000.00
        p_opx, c_opx = -16450000.00, -9850000.00
        p_ebit, c_ebit = 23250000.00, 12850000.00
        p_fin, c_fin = -650000.00, -350000.00
        p_pbt, c_pbt = 22600000.00, 12500000.00
        p_tax, c_tax = -7700130.00, -2985225.00
        p_pat, c_pat = 14899870.00, 9514775.00
        
        p_ppe, c_ppe = 5724540.00, 3409540.00
        p_dta, c_dta = 0.00, 0.00
        p_nca, c_nca = 5724540.00, 3409540.00
        
        p_inv, c_inv = 24300000.00, 13800000.00
        p_ca, c_ca = 6850000.00, 4150000.00
        p_rec, c_rec = 14360000.00, 8685000.00
        p_csh, c_csh = 8924048.00, 5142678.00
        p_cur_a, c_cur_a = 54434048.00, 31777678.00
        p_tot_a, c_tot_a = 60158588.00, 35187218.00
        
        p_sc, c_sc = 1000000.00, 1000000.00
        p_re, c_re = 32376588.00, 17476718.00
        p_eq, c_eq = 33376588.00, 18476718.00
        
        p_bor, c_bor = 7500000.00, 5800000.00
        p_dtl, c_dtl = 98500.00, 16000.00
        p_ncl, c_ncl = 7598500.00, 5816000.00
        
        p_pay, c_pay = 7983500.00, 4825500.00
        p_cl, c_cl = 5800000.00, 3600000.00
        p_ctx, c_ctx = 5400000.00, 2469000.00
        p_cur_l, c_cur_l = 19183500.00, 10894500.00
        p_tot_l_eq, c_tot_l_eq = 60158588.00, 35187218.00
        
        cit_rate_str = "30% (Large Company)"
        tet_rate_str = "3% (Finance Act 2023)"

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    
    # Custom NumberedCanvas configuration
    canvas_maker = NumberedCanvas
    
    # Styles
    styles = getSampleStyleSheet()
    
    c_primary = colors.HexColor("#1F4E79")
    c_secondary = colors.HexColor("#2F5597")
    c_accent = colors.HexColor("#D9E1F2")
    c_dark = colors.HexColor("#262626")
    c_light = colors.HexColor("#F9FAFB")
    c_zebra = colors.HexColor("#F2F4F8")
    
    title_style = ParagraphStyle('CoverTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=20, leading=24, textColor=c_primary, alignment=1)
    subtitle_style = ParagraphStyle('CoverSubtitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, leading=16, textColor=c_secondary, alignment=1)
    h1_style = ParagraphStyle('SectionH1', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=13, leading=16, textColor=c_primary, spaceBefore=8, spaceAfter=4)
    h2_style = ParagraphStyle('SectionH2', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13, textColor=c_secondary, spaceBefore=6, spaceAfter=3)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=11.5, textColor=c_dark)
    body_bold = ParagraphStyle('BodyBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=11.5, textColor=c_dark)
    body_italic = ParagraphStyle('BodyItalic', parent=styles['Normal'], fontName='Helvetica-Oblique', fontSize=8, leading=11, textColor=colors.HexColor("#595959"))
    
    tbl_hdr = ParagraphStyle('TblHdr', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=1)
    tbl_cell_l = ParagraphStyle('TblCellL', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, textColor=c_dark, alignment=0)
    tbl_cell_r = ParagraphStyle('TblCellR', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, textColor=c_dark, alignment=2)
    tbl_cell_c = ParagraphStyle('TblCellC', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, textColor=c_dark, alignment=1)
    tbl_cell_bold_l = ParagraphStyle('TblCellBoldL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=c_dark, alignment=0)
    tbl_cell_bold_r = ParagraphStyle('TblCellBoldR', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=c_dark, alignment=2)

    story = []
    
    # ------------------ COVER PAGE ------------------
    story.append(Spacer(1, 40))
    story.append(Paragraph("BAAY PROJECTS LIMITED", title_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph("(Incorporated in the Federal Republic of Nigeria • RC 1526224 • TIN 21548976-0001)", subtitle_style))
    story.append(Spacer(1, 15))
    story.append(HRFlowable(width="100%", thickness=3, color=c_primary, spaceBefore=5, spaceAfter=20))
    
    story.append(Spacer(1, 40))
    story.append(Paragraph("ANNUAL REPORT AND AUDITED FINANCIAL STATEMENTS", title_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"FOR THE YEAR ENDED 31 DECEMBER {year}", subtitle_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph(f"(With Comparative Figures for the Year Ended 31 December {comp_year})", body_italic))
    story.append(Spacer(1, 20))
    
    # Badge
    badge_data = [[Paragraph("<b>AUDITED FINANCIAL STATEMENTS — IFRS REPORTING PACK</b>", tbl_hdr)]]
    t_badge = Table(badge_data, colWidths=[400])
    t_badge.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), c_primary),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_badge)
    
    story.append(Spacer(1, 80))
    
    # Corporate Information Box
    corp_info = [
        [Paragraph("<b>Registered Corporate Office:</b>", body_bold), Paragraph("Plot 12, Adeola Odeku Street, Victoria Island, Lagos, Nigeria", body_style)],
        [Paragraph("<b>Operational / Site Office:</b>", body_bold), Paragraph("Suite 4B, Admiralty Way, Lekki Phase 1, Lagos, Nigeria", body_style)],
        [Paragraph("<b>Directors:</b>", body_bold), Paragraph("Engr. Babatunde Goke (MD/CEO)<br/>Mrs. Adeola Goke (Executive Director)", body_style)],
        [Paragraph("<b>Company Secretary:</b>", body_bold), Paragraph("Lex Prime Corporate Services, Lagos, Nigeria", body_style)],
        [Paragraph("<b>Independent Auditors:</b>", body_bold), Paragraph("PKF & Co. (Chartered Accountants) / Baker & Associates, Lagos, Nigeria", body_style)],
        [Paragraph("<b>Principal Bankers:</b>", body_bold), Paragraph("Providus Bank Plc • First Bank of Nigeria Limited • Sterling Bank Plc • GTBank", body_style)],
        [Paragraph("<b>Accounting Framework:</b>", body_bold), Paragraph("Full International Financial Reporting Standards (IFRS) & CAMA 2020", body_style)],
    ]
    t_corp = Table(corp_info, colWidths=[150, 360])
    t_corp.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), c_light),
        ('BOX', (0,0), (-1,-1), 1, c_secondary),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E0E0E0")),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_corp)
    
    story.append(PageBreak())
    
    # ------------------ DIRECTORS' REPORT ------------------
    story.append(Paragraph("REPORT OF THE DIRECTORS", h1_style))
    story.append(Paragraph("FOR THE YEAR ENDED 31 DECEMBER " + str(year), h2_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=2, spaceAfter=8))
    
    dir_text = f"""The Directors have pleasure in submitting their report together with the audited financial statements of <b>BAAY PROJECTS LIMITED</b> ("the Company") for the year ended 31 December {year}.<br/><br/>
    <b>1. Principal Activities & Corporate Status:</b><br/>
    The Company is engaged in building construction, civil engineering, land purchase and subdivision for resale, joint venture real estate development, and project management and engineering consultancy. Management has explicitly confirmed that <b>Baay Gokes</b> and <b>Baay Degok</b> are operating trade aliases of Baay Projects Limited sharing the corporate registration number <b>RC 1526224</b>. All operations and accounts are unified within this single legal entity.<br/><br/>
    <b>2. Operating Results:</b><br/>
    The Company recorded statutory gross revenue of <b>₦{p_rev:,.2f}</b> (FY {comp_year}: ₦{c_rev:,.2f}), representing a significant growth driven by milestone delivery in civil engineering infrastructure and serviced land sales. Profit before taxation stood at <b>₦{p_pbt:,.2f}</b> (FY {comp_year}: ₦{c_pbt:,.2f}), while profit for the year after total tax expense of ₦{abs(p_tax):,.2f} was <b>₦{p_pat:,.2f}</b> (FY {comp_year}: ₦{c_pat:,.2f}).<br/><br/>
    <b>3. Dividend:</b><br/>
    The Directors do not recommend the payment of a dividend for the year ended 31 December {year} (FY {comp_year}: Nil). The entire profit for the year has been transferred to retained earnings to fund ongoing development inventories and capital equipment acquisitions.<br/><br/>
    <b>4. Directors & Shareholding:</b><br/>
    The Directors who held office during the year and their beneficial interests in the issued share capital of 1,000,000 Ordinary Shares of ₦1.00 each were as follows:<br/>
    • <b>Engr. Babatunde Goke</b> — 700,000 Ordinary Shares (70.0%)<br/>
    • <b>Mrs. Adeola Goke</b> — 300,000 Ordinary Shares (30.0%)<br/><br/>
    <b>5. Employment of Physically Challenged Persons & Employee Welfare:</b><br/>
    The Company maintains a policy of non-discrimination in recruitment and provides comprehensive on-site occupational safety training, protective equipment, and healthcare allowances for all site and administrative personnel.<br/><br/>
    <b>6. Independent Auditors:</b><br/>
    Messrs. <b>PKF & Co. (Chartered Accountants) / Baker & Associates</b>, having indicated their willingness, will continue in office as statutory auditors in accordance with Section 401(2) of the Companies and Allied Matters Act (CAMA) 2020."""
    story.append(Paragraph(dir_text, body_style))
    story.append(Spacer(1, 15))
    
    # Signature block for Directors
    sig_data = [
        [Paragraph("<b>BY ORDER OF THE BOARD</b>", body_bold), Paragraph("<b>FOR AND ON BEHALF OF THE BOARD</b>", body_bold)],
        [Spacer(1, 20), Spacer(1, 20)],
        [Paragraph("____________________________<br/><b>Lex Prime Corporate Services</b><br/>Company Secretary<br/>Lagos, Nigeria", body_style),
         Paragraph("____________________________<br/><b>Engr. Babatunde Goke</b><br/>Managing Director / CEO<br/>FRC/2014/COREN/00000008912", body_style)]
    ]
    t_sig = Table(sig_data, colWidths=[250, 260])
    t_sig.setStyle(TableStyle([
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(t_sig)
    
    story.append(PageBreak())
    
    # ------------------ STATEMENT OF DIRECTORS' RESPONSIBILITIES ------------------
    story.append(Paragraph("STATEMENT OF DIRECTORS' RESPONSIBILITIES", h1_style))
    story.append(Paragraph("IN RELATION TO THE FINANCIAL STATEMENTS FOR THE YEAR ENDED 31 DECEMBER " + str(year), h2_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=2, spaceAfter=8))
    
    resp_text = f"""The Companies and Allied Matters Act (CAMA) 2020 and the Financial Reporting Council of Nigeria (FRCN) Act require the Directors to prepare financial statements for each financial year that give a true and fair view of the state of financial affairs of <b>BAAY PROJECTS LIMITED</b> at the end of the year and of its financial performance and cash flows for the year then ended.<br/><br/>
    In preparing these financial statements, the Directors are required to:<br/>
    • Select suitable accounting policies in compliance with <b>International Financial Reporting Standards (IFRS)</b> and apply them consistently;<br/>
    • Make judgments and accounting estimates that are reasonable and prudent;<br/>
    • State whether applicable IFRS standards have been followed, subject to any material departures disclosed and explained in the financial statements;<br/>
    • Prepare the financial statements on the going concern basis unless it is inappropriate to presume that the Company will continue in business.<br/><br/>
    The Directors are responsible for keeping proper accounting records that disclose with reasonable accuracy at any time the financial position of the Company and enable them to ensure that the financial statements comply with CAMA 2020 and full IFRS standards. They are also responsible for safeguarding the assets of the Company and taking reasonable steps for the prevention and detection of fraud and other irregularities.<br/><br/>
    The Directors confirm that they have complied with the above requirements in preparing the accompanying draft annual financial statements for the year ended 31 December {year}."""
    story.append(Paragraph(resp_text, body_style))
    story.append(Spacer(1, 25))
    
    dir_sig_box = [
        [Paragraph("____________________________<br/><b>Engr. Babatunde Goke</b><br/>Managing Director / CEO<br/>FRC/2014/COREN/00000008912", body_style),
         Paragraph("____________________________<br/><b>Mrs. Adeola Goke</b><br/>Executive Director (Finance)<br/>FRC/2016/ICAN/00000007841", body_style)]
    ]
    t_dir_sig = Table(dir_sig_box, colWidths=[250, 260])
    story.append(t_dir_sig)
    
    story.append(PageBreak())
    
    # ------------------ INDEPENDENT AUDITOR'S REPORT ------------------
    story.append(Paragraph("INDEPENDENT AUDITOR'S REPORT", h1_style))
    story.append(Paragraph("TO THE MEMBERS OF BAAY PROJECTS LIMITED (RC 1526224)", h2_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=2, spaceAfter=8))
    
    aud_text = f"""<b>Opinion:</b><br/>
    We have audited the financial statements of <b>BAAY PROJECTS LIMITED</b> ("the Company"), which comprise the Statement of Financial Position as at 31 December {year}, the Statement of Profit or Loss and Other Comprehensive Income, Statement of Changes in Equity, and Statement of Cash Flows for the year then ended, and Notes to the Financial Statements, including a summary of significant accounting policies.<br/><br/>
    In our opinion, the accompanying financial statements give a true and fair view of the financial position of Baay Projects Limited as at 31 December {year}, and of its financial performance and its cash flows for the year then ended in accordance with <b>International Financial Reporting Standards (IFRS)</b> and the requirements of the Companies and Allied Matters Act (CAMA) 2020 and the Financial Reporting Council of Nigeria Act.<br/><br/>
    <b>Basis for Opinion:</b><br/>
    We conducted our audit in accordance with International Standards on Auditing (ISAs). Our responsibilities under those standards are further described in the Auditor's Responsibilities for the Audit of the Financial Statements section of our report. We are independent of the Company in accordance with the International Ethics Standards Board for Accountants' International Code of Ethics for Professional Accountants (IESBA Code), and we have fulfilled our other ethical responsibilities. We believe that the audit evidence we have obtained is sufficient and appropriate to provide a basis for our audit opinion.<br/><br/>
    <b>Key Audit Matters (KAMs):</b><br/>
    • <b>IFRS 15 Revenue Recognition & Project WIP:</b> The Company recognizes revenue from civil engineering contracts over time using the cost-to-cost input method. We audited the total contract values, certified progress billings, subcontractor costs incurred, and assessed the reasonableness of forecast costs to complete.<br/>
    • <b>IAS 2 Development Inventories Valuation:</b> Land parcels held for subdivision and ongoing housing construction are stated at the lower of cost and net realizable value (NRV). We reviewed title documentation, purchase deeds, and tested NRV against prevailing real estate market benchmarks.<br/><br/>
    <b>Report on Other Legal and Regulatory Requirements:</b><br/>
    In accordance with the Fifth Schedule of CAMA 2020, we confirm that:<br/>
    i) We have obtained all the information and explanations which to the best of our knowledge and belief were necessary for the purpose of our audit;<br/>
    ii) Proper books of account have been kept by the Company, so far as appears from our examination of those books;<br/>
    iii) The Company's statement of financial position and statement of profit or loss and other comprehensive income are in agreement with the books of account."""
    story.append(Paragraph(aud_text, body_style))
    story.append(Spacer(1, 15))
    
    aud_sig_box = [
        [Paragraph("<b>PKF & Co. (Chartered Accountants) / Baker & Associates</b><br/>Lagos, Nigeria", body_bold)],
        [Spacer(1, 10)],
        [Paragraph("____________________________________________<br/><b>Babajide Cole, FCA</b><br/>Engagement Partner<br/>FRC/2013/ICAN/00000004521<br/>Date: 28 April " + str(year + 1), body_style)]
    ]
    t_aud_sig = Table(aud_sig_box, colWidths=[510])
    story.append(t_aud_sig)
    
    story.append(PageBreak())
    
    # ------------------ PRIMARY STATEMENTS ------------------
    # 1. Statement of Financial Position
    story.append(Paragraph("STATEMENT OF FINANCIAL POSITION", h1_style))
    story.append(Paragraph(f"AS AT 31 DECEMBER {year}", h2_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=2, spaceAfter=6))
    
    sfp_table_data = [
        [Paragraph("<b>Financial Statement Line Item</b>", tbl_hdr), Paragraph("<b>Notes</b>", tbl_hdr), Paragraph(f"<b>31 Dec {year} (₦)</b>", tbl_hdr), Paragraph(f"<b>31 Dec {comp_year} (₦)</b>", tbl_hdr)],
        [Paragraph("<b>NON-CURRENT ASSETS</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph("", tbl_cell_r), Paragraph("", tbl_cell_r)],
        [Paragraph("Property, plant and equipment", tbl_cell_l), Paragraph("7", tbl_cell_c), Paragraph(f"{p_ppe:,.2f}", tbl_cell_r), Paragraph(f"{c_ppe:,.2f}", tbl_cell_r)],
        [Paragraph("Deferred tax asset", tbl_cell_l), Paragraph("15", tbl_cell_c), Paragraph(f"{p_dta:,.2f}", tbl_cell_r), Paragraph(f"{c_dta:,.2f}", tbl_cell_r)],
        [Paragraph("<b>TOTAL NON-CURRENT ASSETS</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_nca:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_nca:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph("<b>CURRENT ASSETS</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph("", tbl_cell_r), Paragraph("", tbl_cell_r)],
        [Paragraph("Inventories (Development land & WIP)", tbl_cell_l), Paragraph("8", tbl_cell_c), Paragraph(f"{p_inv:,.2f}", tbl_cell_r), Paragraph(f"{c_inv:,.2f}", tbl_cell_r)],
        [Paragraph("Contract assets (Unbilled revenue)", tbl_cell_l), Paragraph("9", tbl_cell_c), Paragraph(f"{p_ca:,.2f}", tbl_cell_r), Paragraph(f"{c_ca:,.2f}", tbl_cell_r)],
        [Paragraph("Trade and other receivables", tbl_cell_l), Paragraph("10", tbl_cell_c), Paragraph(f"{p_rec:,.2f}", tbl_cell_r), Paragraph(f"{c_rec:,.2f}", tbl_cell_r)],
        [Paragraph("Cash and cash equivalents", tbl_cell_l), Paragraph("11", tbl_cell_c), Paragraph(f"{p_csh:,.2f}", tbl_cell_r), Paragraph(f"{c_csh:,.2f}", tbl_cell_r)],
        [Paragraph("<b>TOTAL CURRENT ASSETS</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_cur_a:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_cur_a:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph("<b>TOTAL ASSETS</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_tot_a:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_tot_a:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph("<b>EQUITY AND LIABILITIES</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph("", tbl_cell_r), Paragraph("", tbl_cell_r)],
        [Paragraph("Share capital", tbl_cell_l), Paragraph("18", tbl_cell_c), Paragraph(f"{p_sc:,.2f}", tbl_cell_r), Paragraph(f"{c_sc:,.2f}", tbl_cell_r)],
        [Paragraph("Retained earnings", tbl_cell_l), Paragraph("19", tbl_cell_c), Paragraph(f"{p_re:,.2f}", tbl_cell_r), Paragraph(f"{c_re:,.2f}", tbl_cell_r)],
        [Paragraph("<b>TOTAL SHAREHOLDERS' EQUITY</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_eq:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_eq:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph("<b>NON-CURRENT LIABILITIES</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph("", tbl_cell_r), Paragraph("", tbl_cell_r)],
        [Paragraph("Borrowings & director's loans", tbl_cell_l), Paragraph("17", tbl_cell_c), Paragraph(f"{p_bor:,.2f}", tbl_cell_r), Paragraph(f"{c_bor:,.2f}", tbl_cell_r)],
        [Paragraph("Deferred tax liability", tbl_cell_l), Paragraph("15", tbl_cell_c), Paragraph(f"{p_dtl:,.2f}", tbl_cell_r), Paragraph(f"{c_dtl:,.2f}", tbl_cell_r)],
        [Paragraph("<b>TOTAL NON-CURRENT LIABILITIES</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_ncl:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_ncl:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph("<b>CURRENT LIABILITIES</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph("", tbl_cell_r), Paragraph("", tbl_cell_r)],
        [Paragraph("Trade and other payables", tbl_cell_l), Paragraph("14", tbl_cell_c), Paragraph(f"{p_pay:,.2f}", tbl_cell_r), Paragraph(f"{c_pay:,.2f}", tbl_cell_r)],
        [Paragraph("Contract liabilities (Customer advances)", tbl_cell_l), Paragraph("13", tbl_cell_c), Paragraph(f"{p_cl:,.2f}", tbl_cell_r), Paragraph(f"{c_cl:,.2f}", tbl_cell_r)],
        [Paragraph("Current tax liabilities", tbl_cell_l), Paragraph("16", tbl_cell_c), Paragraph(f"{p_ctx:,.2f}", tbl_cell_r), Paragraph(f"{c_ctx:,.2f}", tbl_cell_r)],
        [Paragraph("<b>TOTAL CURRENT LIABILITIES</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_cur_l:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_cur_l:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph("<b>TOTAL EQUITY AND LIABILITIES</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_tot_l_eq:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_tot_l_eq:,.2f}</b>", tbl_cell_bold_r)],
    ]
    t_sfp = Table(sfp_table_data, colWidths=[250, 50, 105, 105])
    t_sfp.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0,0), (-1,-1), 2.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
        ('BACKGROUND', (0,4), (-1,4), c_accent),
        ('BACKGROUND', (0,10), (-1,10), c_accent),
        ('BACKGROUND', (0,11), (-1,11), c_zebra),
        ('BACKGROUND', (0,15), (-1,15), c_accent),
        ('BACKGROUND', (0,19), (-1,19), c_accent),
        ('BACKGROUND', (0,24), (-1,24), c_accent),
        ('BACKGROUND', (0,25), (-1,25), c_zebra),
    ]))
    story.append(t_sfp)
    
    story.append(PageBreak())
    
    # 2. Statement of Profit or Loss
    story.append(Paragraph("STATEMENT OF PROFIT OR LOSS AND OTHER COMPREHENSIVE INCOME", h1_style))
    story.append(Paragraph(f"FOR THE YEAR ENDED 31 DECEMBER {year}", h2_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=2, spaceAfter=6))
    
    pnl_table_data = [
        [Paragraph("<b>Statement Line Item</b>", tbl_hdr), Paragraph("<b>Notes</b>", tbl_hdr), Paragraph(f"<b>Year {year} (₦)</b>", tbl_hdr), Paragraph(f"<b>Year {comp_year} (₦)</b>", tbl_hdr)],
        [Paragraph("Revenue from contracts with customers", tbl_cell_l), Paragraph("3", tbl_cell_c), Paragraph(f"{p_rev:,.2f}", tbl_cell_r), Paragraph(f"{c_rev:,.2f}", tbl_cell_r)],
        [Paragraph("Cost of sales", tbl_cell_l), Paragraph("4", tbl_cell_c), Paragraph(f"({abs(p_cos):,.2f})", tbl_cell_r), Paragraph(f"({abs(c_cos):,.2f})", tbl_cell_r)],
        [Paragraph("<b>GROSS PROFIT</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_gp:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_gp:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph("Administrative and operating expenses", tbl_cell_l), Paragraph("5", tbl_cell_c), Paragraph(f"({abs(p_opx):,.2f})", tbl_cell_r), Paragraph(f"({abs(c_opx):,.2f})", tbl_cell_r)],
        [Paragraph("<b>OPERATING PROFIT (EBIT)</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_ebit:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_ebit:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph("Finance costs", tbl_cell_l), Paragraph("6", tbl_cell_c), Paragraph(f"({abs(p_fin):,.2f})", tbl_cell_r), Paragraph(f"{c_fin:,.2f}", tbl_cell_r)],
        [Paragraph("<b>PROFIT BEFORE TAXATION</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_pbt:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_pbt:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph("Income tax expense", tbl_cell_l), Paragraph("16", tbl_cell_c), Paragraph(f"({abs(p_tax):,.2f})", tbl_cell_r), Paragraph(f"{c_tax:,.2f}", tbl_cell_r)],
        [Paragraph("<b>PROFIT FOR THE YEAR (PAT)</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_pat:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_pat:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph("Other comprehensive income, net of tax", tbl_cell_l), Paragraph("", tbl_cell_c), Paragraph("0.00", tbl_cell_r), Paragraph("0.00", tbl_cell_r)],
        [Paragraph("<b>TOTAL COMPREHENSIVE INCOME</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{p_pat:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_pat:,.2f}</b>", tbl_cell_bold_r)],
    ]
    t_pnl = Table(pnl_table_data, colWidths=[250, 50, 105, 105])
    t_pnl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('BACKGROUND', (0,3), (-1,3), c_accent),
        ('BACKGROUND', (0,5), (-1,5), c_accent),
        ('BACKGROUND', (0,7), (-1,7), c_accent),
        ('BACKGROUND', (0,9), (-1,9), c_zebra),
        ('BACKGROUND', (0,11), (-1,11), c_accent),
    ]))
    story.append(t_pnl)
    story.append(Spacer(1, 15))
    
    # 3. Statement of Changes in Equity
    story.append(Paragraph("STATEMENT OF CHANGES IN EQUITY", h1_style))
    story.append(Paragraph(f"FOR THE YEAR ENDED 31 DECEMBER {year}", h2_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=2, spaceAfter=6))
    
    soce_data = [
        [Paragraph("<b>Equity Item</b>", tbl_hdr), Paragraph("<b>Share Capital (₦)</b>", tbl_hdr), Paragraph("<b>Retained Earnings (₦)</b>", tbl_hdr), Paragraph("<b>Total Equity (₦)</b>", tbl_hdr)],
        [Paragraph(f"<b>Balance at 1 January {comp_year}</b>", tbl_cell_bold_l), Paragraph("1,000,000.00", tbl_cell_r), Paragraph("1,650,000.00", tbl_cell_r), Paragraph("2,650,000.00", tbl_cell_r)],
        [Paragraph(f"Profit for the year {comp_year}", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph(f"{c_pat:,.2f}", tbl_cell_r), Paragraph(f"{c_pat:,.2f}", tbl_cell_r)],
        [Paragraph(f"<b>Balance at 31 December {comp_year}</b>", tbl_cell_bold_l), Paragraph(f"<b>{c_sc:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_re:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_eq:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph(f"Profit for the year {year}", tbl_cell_l), Paragraph("0.00", tbl_cell_r), Paragraph(f"{p_pat:,.2f}", tbl_cell_r), Paragraph(f"{p_pat:,.2f}", tbl_cell_r)],
        [Paragraph(f"<b>Balance at 31 December {year}</b>", tbl_cell_bold_l), Paragraph(f"<b>{p_sc:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{p_re:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{p_eq:,.2f}</b>", tbl_cell_bold_r)],
    ]
    t_soce = Table(soce_data, colWidths=[180, 110, 110, 110])
    t_soce.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('BACKGROUND', (0,3), (-1,3), c_accent),
        ('BACKGROUND', (0,5), (-1,5), c_zebra),
    ]))
    story.append(t_soce)
    
    story.append(PageBreak())
    
    # 4. Statement of Cash Flows
    story.append(Paragraph("STATEMENT OF CASH FLOWS (INDIRECT METHOD)", h1_style))
    story.append(Paragraph(f"FOR THE YEAR ENDED 31 DECEMBER {year}", h2_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=2, spaceAfter=6))
    
    # Calculate CF items dynamically
    dep_val = 385000.00 if year == 2023 else (715000.00 if year == 2024 else 1285000.00)
    ecl_val = 185000.00 if year == 2023 else (260000.00 if year == 2024 else 395000.00)
    fin_val = 125000.00 if year == 2023 else (350000.00 if year == 2024 else 650000.00)
    
    d_inv = -(p_inv - c_inv)
    d_ca = -(p_ca - c_ca)
    d_rec = -(p_rec - c_rec)
    d_pay = p_pay - c_pay
    d_cl = p_cl - c_cl
    
    op_prof_bwc = p_pbt + dep_val + ecl_val + fin_val
    cash_gen_ops = op_prof_bwc + d_inv + d_ca + d_rec + d_pay + d_cl
    tax_paid = 0.00 if year == 2023 else (-2200000.00 if year == 2024 else -4886630.00)
    int_paid = -fin_val
    net_cf_ops = cash_gen_ops + tax_paid + int_paid
    
    ppe_add = -(p_ppe - c_ppe + dep_val)
    net_cf_inv = ppe_add
    
    dir_loan_add = p_bor - c_bor
    net_cf_fin = dir_loan_add
    
    net_csh_delta = net_cf_ops + net_cf_inv + net_cf_fin
    
    cf_table_data = [
        [Paragraph("<b>Cash Flow Operating & Investing Activity</b>", tbl_hdr), Paragraph("<b>Notes</b>", tbl_hdr), Paragraph(f"<b>Year {year} (₦)</b>", tbl_hdr), Paragraph(f"<b>Year {comp_year} (₦)</b>", tbl_hdr)],
        [Paragraph("<b>CASH FLOWS FROM OPERATING ACTIVITIES</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph("", tbl_cell_r), Paragraph("", tbl_cell_r)],
        [Paragraph("Profit before taxation", tbl_cell_l), Paragraph("SPLOCI", tbl_cell_c), Paragraph(f"{p_pbt:,.2f}", tbl_cell_r), Paragraph(f"{c_pbt:,.2f}", tbl_cell_r)],
        [Paragraph("Depreciation of property, plant & equipment", tbl_cell_l), Paragraph("7", tbl_cell_c), Paragraph(f"{dep_val:,.2f}", tbl_cell_r), Paragraph("4,770.00", tbl_cell_r)],
        [Paragraph("Impairment allowance on receivables (ECL)", tbl_cell_l), Paragraph("10", tbl_cell_c), Paragraph(f"{ecl_val:,.2f}", tbl_cell_r), Paragraph("0.00", tbl_cell_r)],
        [Paragraph("Finance costs recognized in profit or loss", tbl_cell_l), Paragraph("6", tbl_cell_c), Paragraph(f"{fin_val:,.2f}", tbl_cell_r), Paragraph("0.00", tbl_cell_r)],
        [Paragraph("<b>Operating profit before working capital changes</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{op_prof_bwc:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_pbt+4770:,.2f}</b>", tbl_cell_bold_r)],
        [Paragraph("(Increase) in development inventories and WIP", tbl_cell_l), Paragraph("8", tbl_cell_c), Paragraph(f"({abs(d_inv):,.2f})", tbl_cell_r), Paragraph("0.00", tbl_cell_r)],
        [Paragraph("(Increase) in contract assets", tbl_cell_l), Paragraph("9", tbl_cell_c), Paragraph(f"({abs(d_ca):,.2f})", tbl_cell_r), Paragraph("0.00", tbl_cell_r)],
        [Paragraph("(Increase) in trade and other receivables", tbl_cell_l), Paragraph("10", tbl_cell_c), Paragraph(f"({abs(d_rec):,.2f})", tbl_cell_r), Paragraph("(1,200,000.00)", tbl_cell_r)],
        [Paragraph("Increase in trade and other payables", tbl_cell_l), Paragraph("14", tbl_cell_c), Paragraph(f"{d_pay:,.2f}", tbl_cell_r), Paragraph("25,000.00", tbl_cell_r)],
        [Paragraph("Increase in contract liabilities (Customer advances)", tbl_cell_l), Paragraph("13", tbl_cell_c), Paragraph(f"{d_cl:,.2f}", tbl_cell_r), Paragraph("0.00", tbl_cell_r)],
        [Paragraph("<b>Cash generated from / (used in) operations</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{cash_gen_ops:,.2f}</b>", tbl_cell_bold_r), Paragraph("<b>(458,522.00)</b>", tbl_cell_bold_r)],
        [Paragraph("Income taxes paid", tbl_cell_l), Paragraph("16", tbl_cell_c), Paragraph(f"{tax_paid:,.2f}", tbl_cell_r), Paragraph("0.00", tbl_cell_r)],
        [Paragraph("Interest paid on project facilities", tbl_cell_l), Paragraph("6", tbl_cell_c), Paragraph(f"({abs(int_paid):,.2f})", tbl_cell_r), Paragraph("0.00", tbl_cell_r)],
        [Paragraph("<b>Net cash flows from / (used in) operating activities</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{net_cf_ops:,.2f}</b>", tbl_cell_bold_r), Paragraph("<b>(458,522.00)</b>", tbl_cell_bold_r)],
        [Paragraph("<b>CASH FLOWS FROM INVESTING ACTIVITIES</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph("", tbl_cell_r), Paragraph("", tbl_cell_r)],
        [Paragraph("Purchase of property, plant and equipment", tbl_cell_l), Paragraph("7", tbl_cell_c), Paragraph(f"({abs(net_cf_inv):,.2f})", tbl_cell_r), Paragraph("0.00", tbl_cell_r)],
        [Paragraph("<b>Net cash flows used in investing activities</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{net_cf_inv:,.2f}</b>", tbl_cell_bold_r), Paragraph("<b>0.00</b>", tbl_cell_bold_r)],
        [Paragraph("<b>CASH FLOWS FROM FINANCING ACTIVITIES</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph("", tbl_cell_r), Paragraph("", tbl_cell_r)],
        [Paragraph("Proceeds from director's project loans", tbl_cell_l), Paragraph("17", tbl_cell_c), Paragraph(f"{net_cf_fin:,.2f}", tbl_cell_r), Paragraph("0.00", tbl_cell_r)],
        [Paragraph("<b>Net cash flows from financing activities</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{net_cf_fin:,.2f}</b>", tbl_cell_bold_r), Paragraph("<b>0.00</b>", tbl_cell_bold_r)],
        [Paragraph("<b>Net increase in cash and cash equivalents</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_c), Paragraph(f"<b>{net_csh_delta:,.2f}</b>", tbl_cell_bold_r), Paragraph("<b>37,500.00</b>", tbl_cell_bold_r)],
        [Paragraph("Cash and cash equivalents at 1 January", tbl_cell_l), Paragraph("11", tbl_cell_c), Paragraph(f"{c_csh:,.2f}", tbl_cell_r), Paragraph("15,000.00", tbl_cell_r)],
        [Paragraph("<b>CASH AND CASH EQUIVALENTS AT 31 DECEMBER</b>", tbl_cell_bold_l), Paragraph("11", tbl_cell_c), Paragraph(f"<b>{p_csh:,.2f}</b>", tbl_cell_bold_r), Paragraph(f"<b>{c_csh:,.2f}</b>", tbl_cell_bold_r)],
    ]
    t_cf = Table(cf_table_data, colWidths=[250, 50, 105, 105])
    t_cf.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('BACKGROUND', (0,6), (-1,6), c_accent),
        ('BACKGROUND', (0,12), (-1,12), c_accent),
        ('BACKGROUND', (0,15), (-1,15), c_zebra),
        ('BACKGROUND', (0,18), (-1,18), c_accent),
        ('BACKGROUND', (0,21), (-1,21), c_accent),
        ('BACKGROUND', (0,24), (-1,24), c_zebra),
    ]))
    story.append(t_cf)
    
    story.append(PageBreak())
    
    # ------------------ NOTES TO THE FINANCIAL STATEMENTS ------------------
    story.append(Paragraph("NOTES TO THE FINANCIAL STATEMENTS", h1_style))
    story.append(Paragraph(f"FOR THE YEAR ENDED 31 DECEMBER {year}", h2_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=2, spaceAfter=8))
    
    notes_narratives = [
        ("1. General Information",
         "<b>BAAY PROJECTS LIMITED</b> ('the Company') was incorporated in Nigeria under the Companies and Allied Matters Act as a private limited liability company on 12 September 2018 (RC 1526224). The registered office is situated at Plot 12, Adeola Odeku Street, Victoria Island, Lagos, Nigeria. The Company's principal activities encompass civil engineering, building construction, land purchase and subdivision for resale, joint venture property developments, and project management engineering consulting. Management has explicitly confirmed that <b>Baay Gokes</b> and <b>Baay Degok</b> are operational business aliases sharing RC 1526224."),
        
        ("2. Summary of Significant Accounting Policies",
         f"<b>2.1 Basis of Preparation:</b> The financial statements have been prepared in accordance with full <b>International Financial Reporting Standards (IFRS)</b> issued by the IASB and adopted by the Financial Reporting Council of Nigeria (FRCN), and CAMA 2020. The financial statements are presented in Nigerian Naira (₦).<br/>"
         f"<b>2.2 IFRS 15 Revenue Recognition:</b> Revenue from construction contracts is recognized over time using the input method (cost-to-cost). Revenue from land and completed property sales is recognized at a point in time when legal title is executed and control transfers.<br/>"
         f"<b>2.3 IAS 2 Inventories:</b> Land held for development/resale and construction WIP are stated at the lower of cost and net realizable value (NRV).<br/>"
         f"<b>2.4 IAS 16 Property, Plant and Equipment:</b> Stated at cost less accumulated depreciation. Depreciation is calculated straight-line: Equipment (25%), Motor Vehicles (25%), Plant & Machinery (20%).<br/>"
         f"<b>2.5 IFRS 9 Financial Instruments:</b> Trade receivables are measured at amortized cost less lifetime expected credit losses (ECL).<br/>"
         f"<b>2.6 IAS 12 Income Taxes:</b> Current tax is provided at statutory rates ({cit_rate_str}, {tet_rate_str}, 0.005% PTF). Deferred tax is recognized on temporary differences."),
        
        (f"3. Revenue from Contracts with Customers (₦{p_rev:,.2f})",
         f"• Civil engineering & infrastructure construction contracts: ₦{p_rev*0.60:,.2f}<br/>"
         f"• Land subdivision & residential serviced plots: ₦{p_rev*0.30:,.2f}<br/>"
         f"• Project management & structural engineering consultancy: ₦{p_rev*0.10:,.2f}"),
        
        (f"4. Cost of Sales (₦{abs(p_cos):,.2f})",
         f"• Construction materials, concrete & steel subcontracts: ₦{abs(p_cos)*0.60:,.2f}<br/>"
         f"• Land & development inventory cost released to sales: ₦{abs(p_cos)*0.30:,.2f}<br/>"
         f"• Direct site technical labor & equipment rentals: ₦{abs(p_cos)*0.10:,.2f}"),
        
        (f"5. Administrative and Operating Expenses (₦{abs(p_opx):,.2f})",
         f"• Staff salaries, site allowances & pension: ₦{abs(p_opx)*0.43:,.2f}<br/>"
         f"• Site office rent, facilities & maintenance: ₦{abs(p_opx)*0.15:,.2f}<br/>"
         f"• Professional, engineering design & legal fees: ₦{abs(p_opx)*0.10:,.2f}<br/>"
         f"• Statutory audit & tax advisory fees: ₦{abs(p_opx)*0.05:,.2f}<br/>"
         f"• Depreciation of property, plant & equipment (IAS 16): ₦{dep_val:,.2f}<br/>"
         f"• Impairment allowance on trade receivables (IFRS 9 ECL): ₦{ecl_val:,.2f}<br/>"
         f"• Motor vehicle fuel, transport & general admin: ₦{abs(p_opx)*0.15:,.2f}"),
        
        (f"6. Finance Costs (₦{abs(p_fin):,.2f})",
         f"Interest expense on commercial bank project facility utilized for site construction financing."),
        
        (f"7. Property, Plant and Equipment (Net Book Value: ₦{p_ppe:,.2f})",
         f"Gross carrying amount ₦{p_ppe+dep_val*(year-2022+1):,.2f} less accumulated depreciation of ₦{dep_val*(year-2022+1):,.2f}."),
        
        (f"8. Development Inventories & WIP (₦{p_inv:,.2f})",
         f"• Land held for development & resale (Epe & Ibeju schemes): ₦{p_inv*0.60:,.2f}<br/>"
         f"• Ongoing civil engineering work-in-progress (Lekki & VI projects): ₦{p_inv*0.40:,.2f}<br/>"
         f"All inventories were tested for impairment; net realizable values substantially exceed carrying costs."),
        
        (f"9. Contract Assets (₦{p_ca:,.2f}) & Contract Liabilities (₦{p_cl:,.2f})",
         f"Contract assets represent unbilled revenue certified for performance satisfied to date. Contract liabilities represent mobilization deposits from off-plan clients."),
        
        (f"10. Trade and Other Receivables (₦{p_rec:,.2f})",
         f"• Gross trade receivables: ₦{p_rec*0.75:,.2f} less IFRS 9 ECL provision ₦{ecl_val*(year-2022):,.2f}<br/>"
         f"• Withholding Tax (WHT) credit notes receivable from clients: ₦{p_rec*0.20:,.2f}<br/>"
         f"• Prepayments and advance payments to materials suppliers: ₦{p_rec*0.05:,.2f}"),
        
        (f"11. Cash and Cash Equivalents (₦{p_csh:,.2f})",
         f"• Providus Bank Plc (A/C 5400281942): ₦{p_csh*0.72:,.2f}<br/>"
         f"• First Bank of Nigeria Limited (A/C 2034891102): ₦{p_csh*0.20:,.2f}<br/>"
         f"• Sterling Bank Plc (A/C 0078451290): ₦{p_csh*0.08:,.2f}<br/>"
         f"All continuous bank statements have been reconciled to the general ledger with zero unexplained differences."),
        
        (f"16. Taxation (Current Tax Charge: ₦{abs(p_tax):,.2f})",
         f"Current tax comprises Companies Income Tax computed under CITA, Tertiary Education Tax (3%) under TETFA, and Police Trust Fund levy (0.005%)."),
        
        (f"17. Related Party Disclosures (IAS 24) (₦{p_bor:,.2f})",
         f"The balance represents unsecured, non-interest bearing project funding advanced by the Managing Director, Engr. Babatunde Goke."),
        
        (f"18. Share Capital (₦1,000,000.00)",
         f"Authorized, issued and fully paid: 1,000,000 Ordinary Shares of ₦1.00 each."),
    ]
    
    for note_title, note_body in notes_narratives:
        story.append(Paragraph(note_title, h2_style))
        story.append(Paragraph(note_body, body_style))
        story.append(Spacer(1, 4))
        
    story.append(PageBreak())
    
    # ------------------ VALUE ADDED & 5-YEAR SUMMARY ------------------
    story.append(Paragraph("STATEMENT OF VALUE ADDED", h1_style))
    story.append(Paragraph(f"FOR THE YEAR ENDED 31 DECEMBER {year}", h2_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=2, spaceAfter=6))
    
    bought_in = abs(p_cos) + (abs(p_opx) - (p_opx*0.43) - dep_val - ecl_val)
    val_added = p_rev - bought_in
    emp_val = abs(p_opx) * 0.43
    gov_val = abs(p_tax)
    prov_val = dep_val + ecl_val
    ret_val = p_pat
    
    va_data = [
        [Paragraph("<b>Value Added Line Item</b>", tbl_hdr), Paragraph(f"<b>Year {year} (₦)</b>", tbl_hdr), Paragraph("<b>%</b>", tbl_hdr)],
        [Paragraph("Gross Revenue from Contracts & Sales", tbl_cell_l), Paragraph(f"{p_rev:,.2f}", tbl_cell_r), Paragraph("100.0%", tbl_cell_c)],
        [Paragraph("Bought-in Materials & Services", tbl_cell_l), Paragraph(f"({bought_in:,.2f})", tbl_cell_r), Paragraph(f"{bought_in/p_rev*100:.1f}%", tbl_cell_c)],
        [Paragraph("<b>VALUE ADDED AVAILABLE FOR DISTRIBUTION</b>", tbl_cell_bold_l), Paragraph(f"<b>{val_added:,.2f}</b>", tbl_cell_bold_r), Paragraph("<b>100.0%</b>", tbl_cell_c)],
        [Paragraph("<b>DISTRIBUTION OF VALUE ADDED:</b>", tbl_cell_bold_l), Paragraph("", tbl_cell_r), Paragraph("", tbl_cell_c)],
        [Paragraph("To Employees (Salaries, wages & benefits)", tbl_cell_l), Paragraph(f"{emp_val:,.2f}", tbl_cell_r), Paragraph(f"{emp_val/val_added*100:.1f}%", tbl_cell_c)],
        [Paragraph("To Government (Corporate taxation & levies)", tbl_cell_l), Paragraph(f"{gov_val:,.2f}", tbl_cell_r), Paragraph(f"{gov_val/val_added*100:.1f}%", tbl_cell_c)],
        [Paragraph("To Providers of Capital (Bank interest)", tbl_cell_l), Paragraph(f"{abs(p_fin):,.2f}", tbl_cell_r), Paragraph(f"{abs(p_fin)/val_added*100:.1f}%", tbl_cell_c)],
        [Paragraph("Retained in Business for Expansion (Depreciation & Reserves)", tbl_cell_l), Paragraph(f"{prov_val+ret_val:,.2f}", tbl_cell_r), Paragraph(f"{(prov_val+ret_val)/val_added*100:.1f}%", tbl_cell_c)],
        [Paragraph("<b>TOTAL VALUE DISTRIBUTED</b>", tbl_cell_bold_l), Paragraph(f"<b>{val_added:,.2f}</b>", tbl_cell_bold_r), Paragraph("<b>100.0%</b>", tbl_cell_c)],
    ]
    t_va = Table(va_data, colWidths=[270, 140, 100])
    t_va.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('BACKGROUND', (0,3), (-1,3), c_accent),
        ('BACKGROUND', (0,9), (-1,9), c_accent),
    ]))
    story.append(t_va)
    story.append(Spacer(1, 15))
    
    story.append(Paragraph("FIVE-YEAR FINANCIAL SUMMARY", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_secondary, spaceBefore=2, spaceAfter=6))
    
    sum_data = [
        [Paragraph("<b>Financial Indicator (₦'000)</b>", tbl_hdr), Paragraph(f"<b>{year}</b>", tbl_hdr), Paragraph(f"<b>{year-1}</b>", tbl_hdr), Paragraph(f"<b>{year-2}</b>", tbl_hdr)],
        [Paragraph("Gross Revenue", tbl_cell_l), Paragraph(f"{p_rev/1000:,.0f}", tbl_cell_r), Paragraph(f"{c_rev/1000:,.0f}", tbl_cell_r), Paragraph("4,200", tbl_cell_r)],
        [Paragraph("Profit Before Taxation", tbl_cell_l), Paragraph(f"{p_pbt/1000:,.0f}", tbl_cell_r), Paragraph(f"{c_pbt/1000:,.0f}", tbl_cell_r), Paragraph("350", tbl_cell_r)],
        [Paragraph("Profit After Taxation", tbl_cell_l), Paragraph(f"{p_pat/1000:,.0f}", tbl_cell_r), Paragraph(f"{c_pat/1000:,.0f}", tbl_cell_r), Paragraph("350", tbl_cell_r)],
        [Paragraph("Share Capital", tbl_cell_l), Paragraph("1,000", tbl_cell_r), Paragraph("1,000", tbl_cell_r), Paragraph("1,000", tbl_cell_r)],
        [Paragraph("Shareholders' Funds (Equity)", tbl_cell_l), Paragraph(f"{p_eq/1000:,.0f}", tbl_cell_r), Paragraph(f"{c_eq/1000:,.0f}", tbl_cell_r), Paragraph("1,650", tbl_cell_r)],
        [Paragraph("Total Assets", tbl_cell_l), Paragraph(f"{p_tot_a/1000:,.0f}", tbl_cell_r), Paragraph(f"{c_tot_a/1000:,.0f}", tbl_cell_r), Paragraph("2,100", tbl_cell_r)],
    ]
    t_sum = Table(sum_data, colWidths=[210, 100, 100, 100])
    t_sum.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('BACKGROUND', (0,3), (-1,3), c_accent),
        ('BACKGROUND', (0,5), (-1,5), c_zebra),
    ]))
    story.append(t_sum)
    
    # Build document
    doc.build(story, canvasmaker=canvas_maker)
    print(f"Successfully generated '{output_path}'!")

if __name__ == '__main__':
    create_afs_pdf(2023, "BAAY_PROJECTS_LIMITED_AFS_2023.pdf")
    create_afs_pdf(2024, "BAAY_PROJECTS_LIMITED_AFS_2024.pdf")
    create_afs_pdf(2025, "BAAY_PROJECTS_LIMITED_AFS_2025.pdf")
