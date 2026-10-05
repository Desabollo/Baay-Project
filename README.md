# Baay Project — 2023–2025 audit model

## Versioned deliverables

The original files remain under their original filenames. Revised deliverables are clearly identified by the `_updated` suffix:

- `BAAY_PROJECTS_LIMITED_Master_Workbook_2023_2025_updated.xlsx` — updated 36-sheet integrated master workbook.
- `BAAY_PROJECTS_LIMITED_AFS_2023_updated.pdf`
- `BAAY_PROJECTS_LIMITED_AFS_2024_updated.pdf`
- `BAAY_PROJECTS_LIMITED_AFS_2025_updated.pdf`
- `BAAY_PROJECTS_LIMITED_Working_Papers_and_Ledgers_updated.pdf`
- `BAAY_PROJECTS_LIMITED_Information_Gap_and_Completion_Report_updated.pdf`
- `BAAY_Sales_to_Bank_Reconciliation_Report_updated.xlsx` — detailed annual bridge, bank-account coverage and source drill-down.
- `BAAY_Audit_Integrity_Checks_Report_updated.md` — independent 12-point re-verification output.
- `Status Report- BAAY PROJECTS LTD.pdf` — CAC source document used for the corporate particulars update.

The revised reporting set uses the CAC status report dated 26 August 2026: registered office No. 7 Zika Usifo Street, Ikosi Ketu, Agege, Lagos State; RC 1526224; former name BAAY DEGOK NIG LTD; company secretary Adegoke Mary Ayoboade; the active-director register and current shareholder register; and Sanni Waheed & Co. as external auditors. The engagement partner is shown with FRCN number `FRC/2016/ICAN/2016/00000013886`. Directors are not assigned FRCN numbers.

The AFS cover/header no longer uses “IFRS REPORTING PACK” or “Statutory Working Papers”. The statement of changes in equity is presented on a separate page. Notes use tabular note presentation, accounting policies use the supplied policy/rate format, and operating results use the supplied turnover/profit-forward format. The AFS now reflect the N99,000,000 increase in issued share capital effective in 2023. Because the additional shares were unpaid, the AFS present N100,000,000 issued share capital from 2023 onward less N99,000,000 due from shareholders as an offset to equity, with no asset or cash-flow proceeds recognized. The 2022 comparative remains N1,000,000. This treatment follows Section 22.7(a) of the IFRS for SMEs Accounting Standard.

## Rebuild and verify the updated versions

```bash
python build_master_workbook.py
python audit_integrity_checks.py
python pdf_generator_afs.py
python pdf_generator_working_papers.py
python pdf_generator_completion_report.py
```

`build_master_workbook.py` invokes `integrate_sales_customer_data.py`, parses all eight source tabs and 2,576 source formulas, embeds the evaluated source populations in the existing 36-sheet structure, and produces the detailed reconciliation workbook.

The accountant workbook identifies its figures as cash receipts rather than IFRS 15 revenue. The integration therefore retains source receipts as a separate subledger and does not automatically post the sales-to-bank control difference into revenue or use it as a balancing plug.
