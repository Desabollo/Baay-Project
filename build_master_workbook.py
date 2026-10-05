import os
import sys
import openpyxl

import sheet_builders_1
import sheet_builders_2
import sheet_builders_3
import sheet_builders_4
import sheet_builders_5
import sheet_builders_6
from integrate_sales_customer_data import integrate

print("Assembling BAAY_PROJECTS_LIMITED_Master_Workbook_2023_2025_updated.xlsx...")

wb = openpyxl.Workbook()
wb.remove(wb.active) # Remove default sheet

# Sheets 1 to 8
print("Building Sheets 1 to 8...")
sheet_builders_1.build_sheet_1_control_readme(wb)
sheet_builders_1.build_sheet_2_entity_register(wb)
sheet_builders_1.build_sheet_3_sources_infogap(wb)
sheet_builders_1.build_sheet_4_opening_bridge(wb)
sheet_builders_1.build_sheet_5_chart_of_accounts(wb)
sheet_builders_1.build_sheet_6_coverage_bank_matrix(wb)
sheet_builders_1.build_sheet_7_normalized_bank_data(wb)
sheet_builders_1.build_sheet_8_interbank_transfers(wb)

# Sheets 9 to 14
print("Building Sheets 9 to 14...")
sheet_builders_2.build_sheet_9_journal_entries(wb)
sheet_builders_2.build_sheet_10_gl_ledgers(wb)
sheet_builders_2.build_sheet_11_tb_unadjusted(wb)
sheet_builders_2.build_sheet_12_tb_adjusted_preclosing(wb)
sheet_builders_2.build_sheet_13_tb_postclosing(wb)
sheet_builders_2.build_sheet_14_gl_index_mapping(wb)

# Sheets 15 to 19
print("Building Sheets 15 to 19...")
sheet_builders_3.build_sheet_15_project_register(wb)
sheet_builders_3.build_sheet_16_ifrs15_revenue_wip(wb)
sheet_builders_3.build_sheet_17_ppe_depreciation(wb)
sheet_builders_3.build_sheet_18_receivables_ecl_payables(wb)
sheet_builders_3.build_sheet_19_related_parties_equity(wb)

# Sheets 20 to 26
print("Building Sheets 20 to 26...")
sheet_builders_4.build_sheet_20_tax_law_matrix(wb)
sheet_builders_4.build_sheet_21_cit_tet_mintax(wb)
sheet_builders_4.build_sheet_22_capital_allowances_losses(wb)
sheet_builders_4.build_sheet_23_vat_monthly(wb)
sheet_builders_4.build_sheet_24_wht_receivable_payable(wb)
sheet_builders_4.build_sheet_25_deferred_tax_ias12(wb)
sheet_builders_4.build_sheet_26_tax_payable_rollforward(wb)

# Sheets 27 to 32
print("Building Sheets 27 to 32...")
sheet_builders_5.build_sheet_27_afs_2023(wb)
sheet_builders_5.build_sheet_28_notes_2023(wb)
sheet_builders_5.build_sheet_29_afs_2024(wb)
sheet_builders_5.build_sheet_30_notes_2024(wb)
sheet_builders_5.build_sheet_31_afs_2025(wb)
sheet_builders_5.build_sheet_32_notes_2025(wb)

# Sheets 33 to 36
print("Building Sheets 33 to 36...")
sheet_builders_6.build_sheet_33_three_year_summary(wb)
sheet_builders_6.build_sheet_34_cash_flow_workings(wb)
sheet_builders_6.build_sheet_35_assumptions_estimates(wb)
sheet_builders_6.build_sheet_36_exception_dashboard_checks(wb)

output_filename = "BAAY_PROJECTS_LIMITED_Master_Workbook_2023_2025_updated.xlsx"
wb.save(output_filename)
print(f"Master Workbook saved successfully as '{output_filename}'! Total sheets: {len(wb.sheetnames)}")
print("Integrating accountant sales/customer schedules and creating reconciliation report...")
integrate()
print("Sales/customer integration completed; 36-sheet structure preserved.")
