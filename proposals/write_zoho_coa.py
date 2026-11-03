"""Zoho Books chart of accounts for Baay Projects Limited.

The grain follows the controlled chart in
BAAY_PROJECTS_LIMITED_Master_Workbook_2023_2025_updated.xlsx on the
Accounting & Tax branch: banks apart, receivables and revenue by stream,
inventory in three buckets, tax by levy, and expenses by nature.

Accounts already in the Zoho sample are coded and written in red in the
workbook. A CSV cannot carry colour, so the upload file is the same chart
without the review columns. Parents are listed before their children, and a
child always has the same Zoho account type as its parent.
"""

from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
ROOT = Path(__file__).resolve().parents[1]
OUT_CSV = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Chart_of_Accounts.csv"
OUT_XLSX = ROOT / "BAAY_PROJECTS_LIMITED_Zoho_Books_Chart_of_Accounts.xlsx"
SAMPLE = Path("/tmp/arena-coding-uploads/Zoho Books Chart_of_Accounts.csv")

# Zoho matches an overwrite on the account name. These names are the sample export.
SAMPLE_NAMES = {
    "Other Charges",
    "Lodging",
    "Undeposited Funds",
    "Petty Cash",
    "Accounts Receivable",
    "Furniture and Equipment",
    "Advance Tax",
    "Accounts Payable",
    "VAT Payable",
    "Retained Earnings",
    "Owner's Equity",
    "Opening Balance Offset",
    "Sales",
    "General Income",
    "Unearned Revenue",
    "Interest Income",
    "Late Fee Income",
    "Office Supplies",
    "Advertising And Marketing",
    "Discount",
    "Bank Fees and Charges",
    "Credit Card Charges",
    "Exchange Gain or Loss",
    "Travel Expense",
    "Telephone Expense",
    "Automobile Expense",
    "IT and Internet Expenses",
    "Rent Expense",
    "Janitorial Expense",
    "Shipping Charge",
    "Postage",
    "Bad Debt",
    "Employee Advance",
    "Opening Balance Adjustments",
    "Printing and Stationery",
    "Employee Reimbursements",
    "Uncategorized",
    "Salaries and Employee Wages",
    "Meals and Entertainment",
    "Depreciation Expense",
    "Consultant Expense",
    "Dimension Adjustments",
    "Repairs and Maintenance",
    "Other Expenses",
    "Inventory Asset",
    "Cost of Goods Sold",
    "Drawings",
    "Purchase Discounts",
    "Prepaid Expenses",
}

# Locked or tax-module accounts. A parent on these rows is the usual reason an
# import fails, so the upload file leaves Parent Account blank. The review
# sheet still shows the group they belong in.
LOCKED = {
    "Undeposited Funds",
    "Accounts Receivable",
    "Accounts Payable",
    "VAT Payable",
    "Retained Earnings",
    "Opening Balance Offset",
    "Unearned Revenue",
    "Discount",
    "Late Fee Income",
    "Shipping Charge",
    "Other Charges",
    "Opening Balance Adjustments",
    "Dimension Adjustments",
    "Inventory Asset",
    "Cost of Goods Sold",
    "Sales",
    "Exchange Gain or Loss",
}

# code, name, zoho type, parent name or "", workbook ref, post, do_not
# workbook ref is the code in the master workbook, or "" if this line is a
# header, a Zoho default, or an account added because the books need it.
ROWS = [
    # Cash at bank. Workbook 1010-1040. Transfers are not income.
    ("1000", "Cash at bank", "Bank", "", "",
     "Header only. Do not post.",
     "Do not post a receipt, a payment or a transfer to this header."),
    ("1010", "Providus Bank", "Bank", "Cash at bank", "1010",
     "The Providus operations account in the master workbook. Bank feed or weekly statement import.",
     "A transfer to First Bank, Sterling or GTBank is a transfer, not income and not an expense."),
    ("1020", "First Bank of Nigeria", "Bank", "Cash at bank", "1020",
     "The First Bank collections account in the master workbook.",
     "Do not credit sales because a subscriber's name appears on the narration."),
    ("1030", "Sterling Bank", "Bank", "Cash at bank", "1030",
     "The Sterling disbursements account in the master workbook.",
     "Do not record a movement from Providus as a receipt."),
    ("1040", "Guaranty Trust Bank", "Bank", "Cash at bank", "1040",
     "The GTBank account in the master workbook.",
     "Do not invent a balance for a month whose statement has not been received. Leave the feed or the import blank and say so."),

    ("1080", "Cash on hand", "Cash", "", "",
     "Header only. Do not post.",
     "Do not post bank balances here."),
    ("1050", "Petty Cash", "Cash", "Cash on hand", "1050",
     "Small cash held in the office. Workbook 1050.",
     "Do not use this for a bank balance, and do not use it to hide a payment whose nature is unknown."),
    ("1055", "Site imprest", "Cash", "Cash on hand", "1050",
     "Cash float held at a site. Split out of workbook 1050 so an office float is not mixed with a site float.",
     "When the float is reimbursed, credit this account, not income."),
    ("1090", "Undeposited Funds", "Cash", "", "",
     "Zoho clearing account for a receipt recorded before it is deposited.",
     "Do not leave subscriber money here overnight. Deposit it to the bank that received it."),

    # One Accounts Receivable. Zoho will not allow a second.
    ("1100", "Accounts Receivable", "Accounts Receivable", "", "",
     "Zoho's only customer-invoice control account. An invoice raised in Zoho posts here.",
     "Do not raise a second receivable account of this type. Land, housing and construction are split on 1110, 1120 and 1125 by journal, or by the project on the invoice."),

    ("1105", "Trade receivables", "Other Current Asset", "", "",
     "Header only. The sum is trade receivables on the statement of financial position.",
     "Do not post to the header."),
    ("1110", "Trade receivables - construction and engineering", "Other Current Asset", "Trade receivables", "1110",
     "Unpaid balance of a construction contract whose revenue has been recognised. Workbook 1110.",
     "Do not debit this for a receipt. A receipt before certification is Unearned Revenue."),
    ("1120", "Trade receivables - land and property", "Other Current Asset", "Trade receivables", "1120",
     "Unpaid balance of a plot whose control has transferred, and not above the balance on that subscriber's register. Workbook 1120.",
     "Do not recognise a receivable above the customer register, and do not use this for a subscriber who is still Paying."),
    ("1125", "Trade receivables - housing", "Other Current Asset", "Trade receivables", "",
     "Unpaid balance of a house whose handover has been recorded. Added because the sales register separates housing from land. Same grain as workbook 1120.",
     "No handover, no receivable. A receipt before handover is Unearned Revenue."),
    ("1130", "Allowance for expected credit losses", "Other Current Asset", "Trade receivables", "1130",
     "Credit balance. Five per cent of gross trade receivables, as the company already provides. Workbook 1130.",
     "Do not post the expense here. The expense is Bad Debt, 6080. Do not use this account to write off a named debt without a journal."),

    ("1145", "Other current assets", "Other Current Asset", "", "",
     "Header only.",
     "Do not post to the header."),
    ("1140", "Contract assets", "Other Current Asset", "Other current assets", "1140",
     "Certified work not yet invoiced. Workbook 1140. Use only where a contract has been certified and the invoice has not been raised.",
     "A subscriber instalment is not a contract asset. It is Unearned Revenue until allocation or handover."),
    ("1150", "Withholding tax credit receivable", "Other Current Asset", "Other current assets", "1150",
     "WHT deducted by customers and held as a credit against companies income tax. Workbook 1150.",
     "Do not net this against VAT, and do not take it to income."),
    ("1158", "Advance Tax", "Other Current Asset", "Other current assets", "",
     "Tax paid to the revenue service before the assessment. Zoho sample account, kept for that purpose.",
     "A customer WHT credit note goes to 1150, not here."),
    ("1160", "Prepaid Expenses", "Other Current Asset", "Other current assets", "1160",
     "Rent, insurance and supplier advances paid before the period they belong to. Workbook 1160, prepayments and site advances.",
     "Do not leave a site advance here once the related cost has been incurred. Release it to inventory or to expense."),
    ("1165", "Employee Advance", "Other Current Asset", "Other current assets", "",
     "Cash given to a staff member before the expense is incurred.",
     "Director drawings and dividends do not belong here."),
    ("1170", "Other receivables", "Other Current Asset", "Other current assets", "",
     "Amounts recorded as due, other than trade receivables on contracts whose revenue has been recognised.",
     "Do not use this as a home for a bank line you have not read."),
    ("1175", "Amounts due from related parties", "Other Current Asset", "Other current assets", "",
     "Net of amounts received from and paid to the directors, after director remuneration and after dividends. A debit is due to the company.",
     "Do not mix subscriber money into this account. Do not net it against the director's long-term loan, 2510."),
    ("1180", "Payments not yet identified", "Other Current Asset", "Other current assets", "",
     "A payment whose nature is not yet known. It waits here, outside profit, until it is named.",
     "Do not post it to an expense to make the month look tidy."),
    ("1185", "Foreign-currency and other advances", "Other Current Asset", "Other current assets", "",
     "Foreign-currency balances and other advances that are not a subscriber receipt and not inventory.",
     "Do not take an exchange difference to revenue."),
    ("1190", "Loan repayments in excess of proceeds identified", "Other Current Asset", "Other current assets", "",
     "Repayments recorded in excess of loan proceeds that have been identified. Asset until the proceeds are found.",
     "Do not treat the excess as a finance cost."),

    ("1205", "Development inventory", "Stock", "", "",
     "Header only. Land, work in progress and completed units.",
     "Do not post to the header. Do not use Zoho's Inventory Asset for a plot or a house."),
    ("1210", "Land held for resale", "Stock", "Development inventory", "1210",
     "Cost of development land not yet released to a sale. Workbook 1210.",
     "Release cost to 5020 only when the related land revenue is recognised."),
    ("1220", "Construction work in progress", "Stock", "Development inventory", "1220",
     "Development expenditure on units not yet completed. Workbook 1220.",
     "Do not expense site costs that still belong to a unit."),
    ("1230", "Completed units held for resale", "Stock", "Development inventory", "1230",
     "Completed houses and units not yet handed over. Workbook 1230.",
     "Handover releases this cost to 5025. A receipt is not the handover."),
    ("1290", "Inventory Asset", "Stock", "", "",
     "Zoho's inventory-module account. Leave it for any item the inventory module tracks.",
     "Do not post plot cost, housing cost or development expenditure here."),

    ("1500", "Property plant and equipment", "Fixed Asset", "", "",
     "Header only. Carrying amount is cost less 1590.",
     "Do not post to the header. Do not expense a purchase that will be used for more than a year."),
    ("1510", "Office and site equipment", "Fixed Asset", "Property plant and equipment", "1510",
     "Office equipment, site equipment and computers. Workbook 1510. Depreciation is 20 per cent a year, straight line, a full year in the year of purchase.",
     "A repair that does not extend life goes to Repairs and Maintenance, not here."),
    ("1519", "Furniture and Equipment", "Fixed Asset", "Property plant and equipment", "",
     "Zoho sample account. Do not use it for a new purchase. Post new equipment to 1510, and move any opening balance there.",
     "Do not depreciate this account as well as 1510."),
    ("1520", "Motor and project vehicles", "Fixed Asset", "Property plant and equipment", "1520",
     "Motor vehicles and project vehicles. Workbook 1520. Same depreciation rate.",
     "Fuel and servicing go to Automobile Expense, not here."),
    ("1530", "Plant and machinery", "Fixed Asset", "Property plant and equipment", "1530",
     "Plant and machinery. Workbook 1530. Same depreciation rate.",
     "Hired plant is a cost of sales or a repair, not an asset."),
    ("1590", "Accumulated depreciation", "Fixed Asset", "Property plant and equipment", "1590",
     "Credit balance. Accumulated depreciation of 1510, 1520 and 1530. Workbook 1590.",
     "The charge for the year goes to Depreciation Expense, 6070. Do not debit this account for a purchase."),

    ("1600", "Other non-current assets", "Other Asset", "", "",
     "Header only.",
     "Do not post to the header."),
    ("1610", "Deferred tax asset", "Other Asset", "Other non-current assets", "1610",
     "Deferred tax asset under IAS 12. Workbook 1610. The company has not been recognising deferred tax. Open this only if that changes.",
     "Do not use this for WHT credits or tax paid in advance."),

    ("2000", "Accounts Payable", "Accounts Payable", "", "",
     "Zoho's only supplier-bill control account. A bill raised in Zoho posts here.",
     "Do not create a second account of this type. A journal split, if the face of the statement needs one, goes to 2010."),
    ("2005", "Trade and other payables", "Other Current Liability", "", "",
     "Header only.",
     "Do not post to the header."),
    ("2010", "Trade payables - subcontractors and materials", "Other Current Liability", "Trade and other payables", "2010",
     "Subcontractors and materials not entered as a Zoho bill. Workbook 2010.",
     "A subscriber deposit is not a payable."),
    ("2030", "Accrued audit tax and professional fees", "Other Current Liability", "Trade and other payables", "2030",
     "Audit, tax and professional fees incurred and not yet billed. Workbook 2030.",
     "Do not accrue a fee that has not been agreed."),
    ("2040", "Other accrued expenses and sundry creditors", "Other Current Liability", "Trade and other payables", "2040",
     "Other accruals and sundry creditors. Workbook 2040.",
     "An unidentified bank payment does not belong here. It belongs in 1180."),
    ("2048", "Employee Reimbursements", "Other Current Liability", "Trade and other payables", "",
     "Expenses a staff member has incurred and the company has not yet repaid.",
     "An advance that has not been spent is Employee Advance, an asset."),
    ("2090", "Opening Balance Adjustments", "Other Current Liability", "", "",
     "Zoho uses this for the difference on an opening balance. It should be nil once the opening trial balance ties to the signed statement of financial position at 31 December 2025.",
     "Do not post the year's transactions here."),

    ("2020", "Unearned Revenue", "Other Current Liability", "", "2020",
     "Subscriber receipts before the plot is allocated or the house is handed over. This is workbook 2020, contract liabilities. Zoho also uses this account for retainers.",
     "Do not credit sales from a receipt. Do not put co-ownership here. Do not recognise revenue until Finance records the allocation or the handover."),
    ("2024", "Customer and other liabilities", "Other Current Liability", "", "",
     "Header only. Liabilities that are not trade payables and not tax.",
     "Do not post to the header."),
    ("2025", "Co-ownership liabilities", "Other Current Liability", "Customer and other liabilities", "",
     "Investment receipts that are not a sale of land or housing, less returns paid. Added because those receipts are in the sales records and must not become revenue.",
     "Never credit 4010, 4020 or 4025 from this account."),
    ("2045", "Receipts not yet identified", "Other Current Liability", "Customer and other liabilities", "",
     "A receipt whose nature is not yet known. It waits here, outside profit.",
     "Do not credit income so that a month will balance."),
    ("2055", "Amounts due on the company's other accounts", "Other Current Liability", "Customer and other liabilities", "",
     "Net amount due on accounts of the company other than the four bank balances in cash.",
     "Do not treat a transfer between the company's own banks as this account."),

    ("2100", "Current tax liabilities", "Other Current Liability", "", "",
     "Header only. The charge for each year, less tax payments identified. Not an agreed assessment.",
     "Do not post to the header. Credit the levy that applies."),
    ("2110", "Companies income tax payable", "Other Current Liability", "Current tax liabilities", "2110",
     "CIT payable. Workbook 2110. Credit from 8010.",
     "Do not put TET, the police trust fund levy or NASENI in this account."),
    ("2120", "Tertiary education tax payable", "Other Current Liability", "Current tax liabilities", "2120",
     "Tertiary education tax payable. Workbook 2120.",
     "Do not net this against CIT."),
    ("2130", "Police trust fund levy payable", "Other Current Liability", "Current tax liabilities", "2130",
     "Police trust fund levy payable. Workbook 2130.",
     "Do not omit it because the amount is small."),
    ("2140", "NASENI levy payable", "Other Current Liability", "Current tax liabilities", "2140",
     "NASENI levy payable for a large company. Workbook 2140.",
     "Do not charge it for a year in which the company is not large."),
    ("2150", "VAT Payable", "Other Current Liability", "", "2150",
     "Zoho's VAT account. Output less input. Workbook 2150.",
     "Do not charge VAT on residential land and housing sales. Do not journal VAT that a tax code has already posted."),
    ("2160", "Withholding tax payable", "Other Current Liability", "Current tax liabilities", "2160",
     "WHT deducted from vendors and not yet remitted. Workbook 2160.",
     "Do not expense the deduction. The expense is the gross bill. This account is the amount still to be paid over."),

    ("2500", "Borrowings and related-party funding", "Long Term Liability", "", "",
     "Header only.",
     "Do not post a subscriber receipt here."),
    ("2510", "Director loan and long-term project funding", "Long Term Liability", "Borrowings and related-party funding", "2510",
     "Director's loan and other long-term project funding. Workbook 2510.",
     "A current account with a director that is not a loan goes to 1175. Interest goes to 7010, not a reduction of the loan without a journal."),
    ("2520", "Bank project facility", "Long Term Liability", "Borrowings and related-party funding", "",
     "Drawn bank facility, where the proceeds have been identified. Home for the interest in 7010.",
     "A repayment in excess of proceeds identified goes to 1190, not a debit that drives this account into an unexamined credit."),

    ("2600", "Deferred tax and other liabilities", "Other Liability", "", "",
     "Header only.",
     "Do not post to the header."),
    ("2610", "Deferred tax liability", "Other Liability", "Deferred tax and other liabilities", "2610",
     "Deferred tax liability under IAS 12. Workbook 2610. Open it only if deferred tax is recognised.",
     "Current tax does not belong here."),
    ("9098", "Dimension Adjustments", "Other Liability", "", "",
     "Zoho system account for tag and branch transfers. Leave it alone.",
     "Do not use it for an unidentified bank line."),

    ("3000", "Share capital and other equity", "Equity", "", "",
     "Header only. Retained earnings stays outside this header because Zoho holds that account itself.",
     "Do not post profit here. Profit closes to Retained Earnings."),
    ("3010", "Ordinary share capital", "Equity", "Share capital and other equity", "3010",
     "Issued ordinary shares of 1 naira each, to the extent paid. Workbook 3010. The amount recognised is the paid capital.",
     "Do not credit this account for a subscriber receipt, and do not credit it for shares that have not been paid."),
    ("3012", "Amount due on issued shares", "Equity", "Share capital and other equity", "",
     "Debit balance. Unpaid amount on issued shares, deducted from equity under IAS 32. Not a receivable.",
     "Do not invoice the shareholder through Accounts Receivable for this balance."),
    ("3030", "Dividends", "Equity", "Share capital and other equity", "",
     "Dividends declared and paid. Debit balance. Closed to retained earnings. Not an expense.",
     "Do not post a dividend to Salaries, Director remuneration, or Drawings."),
    ("3019", "Owner's Equity", "Equity", "Share capital and other equity", "",
     "Zoho sample account for an unincorporated owner. Do not use it. This company posts share capital to 3010.",
     "Do not put opening retained earnings here."),
    ("3090", "Drawings", "Equity", "Share capital and other equity", "",
     "Zoho sample account for an unincorporated owner. Do not use it. A dividend is 3030. A director's loan movement is 2510 or 1175.",
     "Do not post subscriber money here."),
    ("3020", "Retained Earnings", "Equity", "", "3020",
     "Zoho's retained earnings account. Workbook 3020. Profit closes here.",
     "Do not journal the year's profit into this account by hand, and do not post dividends straight to it if 3030 is in use. Close 3030 here."),
    ("3098", "Opening Balance Offset", "Equity", "", "",
     "Zoho's opening-balance buffer. It should be nil once the opening trial balance agrees to the signed statement of financial position at 31 December 2025.",
     "Do not use it to force a month to balance."),

    ("4000", "Revenue from contracts with customers", "Income", "", "",
     "Header only. Estates are Zoho projects, not further accounts: Green City Lantaba, Idobi and Omu-Epe, Green City Ibadan, Heritage Garden, Agbowa, Green Hillside, Pacific Court, Pacific Apartment, Gorge View Court and Baay Foreshore.",
     "Do not post to the header. A receipt is not revenue."),
    ("4010", "Revenue - civil engineering and construction", "Income", "Revenue from contracts with customers", "4010",
     "Construction and engineering revenue when the related performance is complete. Workbook 4010.",
     "Do not use this for a plot sale or a housing handover."),
    ("4020", "Revenue - land and property sales", "Income", "Revenue from contracts with customers", "4020",
     "Land revenue when control of an identified plot has transferred, and not above the balance on the customer register. Workbook 4020.",
     "A subscriber who is still Paying has not produced revenue. Credit Unearned Revenue."),
    ("4025", "Revenue - housing", "Income", "Revenue from contracts with customers", "",
     "Housing revenue, and only when handover is recorded. Added because the sales register separates housing from land.",
     "The transfer of money is not the handover. Do not let the software recognise this in the background."),
    ("4030", "Revenue - project management and consulting", "Income", "Revenue from contracts with customers", "4030",
     "Project management and engineering consulting. Workbook 4030.",
     "Do not bury a plot sale in this account."),
    ("4040", "Revenue - ancillary charges", "Income", "Revenue from contracts with customers", "",
     "Corner piece, documentation and development charges. They are income of their own, not a silent addition to the plot price.",
     "Do not net them against cost of sales."),
    ("4099", "Sales", "Income", "", "",
     "Zoho sample income account. Do not use it for estate sales. Point items and journals at 4010 to 4040.",
     "Do not post a bank receipt here."),
    ("4095", "Discount", "Income", "", "",
     "Zoho posts an invoice discount here. Leave the account in place.",
     "A refund of a subscriber receipt is not a discount. It reverses Unearned Revenue or the sale it relates to."),
    ("4096", "Shipping Charge", "Income", "", "",
     "Zoho sample account for a shipping line on an invoice. Not used for land or housing.",
     "Do not post development charges here. Those are 4040."),
    ("4097", "Other Charges", "Income", "", "",
     "Zoho sample account for an adjustment line on an invoice.",
     "Do not use it as a general income account. Sundry receipts are 4150."),
    ("4120", "Late Fee Income", "Income", "", "",
     "Zoho sample account for a late fee raised on an invoice.",
     "A subscriber's next instalment is not a late fee."),

    ("4100", "Other income", "Income", "", "",
     "Header only. Income that is not a sale of land, housing or construction.",
     "Do not post a plot receipt here."),
    ("4110", "Interest Income", "Income", "Other income", "",
     "Interest earned on bank balances.",
     "A bank charge is not negative interest. It goes to Bank Fees and Charges."),
    ("4130", "Other trading income", "Income", "Other income", "",
     "Trading that is not a plot or a house. Kept separate so a trade does not look like a land sale.",
     "The related cost goes to 5040, not to land released."),
    ("4140", "Service income", "Income", "Other income", "",
     "Service income that is not project-management revenue.",
     "Do not use this for co-ownership."),
    ("4150", "Sundry receipts", "Income", "Other income", "",
     "Small receipts that have been identified and are income, and are not a sale.",
     "If the narration has not been identified, use 2045, not this account."),
    ("4190", "General Income", "Income", "Other income", "",
     "Zoho sample account. Use 4130, 4140 or 4150 for a receipt you can name. Use this only if none of those fits.",
     "Do not use it as a dumping account for the bank feed."),

    ("5000", "Cost of sales", "Cost Of Goods Sold", "", "",
     "Header only.",
     "Do not post to the header. An administrative cost does not belong in this group."),
    ("5010", "Cost of sales - materials and subcontracts", "Cost Of Goods Sold", "Cost of sales", "5010",
     "Construction materials and subcontractors released to a contract. Workbook 5010.",
     "Do not use this for the cost of a plot."),
    ("5020", "Cost of sales - land and property released", "Cost Of Goods Sold", "Cost of sales", "5020",
     "Land cost released when the related land revenue is recognised. Workbook 5020.",
     "Do not release cost on a receipt. Release it when 4020 is credited."),
    ("5025", "Cost of sales - housing released", "Cost Of Goods Sold", "Cost of sales", "",
     "Housing cost released when handover is recorded and 4025 is credited.",
     "Do not release the cost of a unit that is still in 1220 or 1230."),
    ("5030", "Cost of sales - site labour and equipment hire", "Cost Of Goods Sold", "Cost of sales", "5030",
     "Direct site labour and hired equipment. Workbook 5030.",
     "Office salaries are 6010, not here."),
    ("5040", "Cost of sales - trading purchases", "Cost Of Goods Sold", "Cost of sales", "",
     "Purchases for the other trading in 4130.",
     "Do not post plot cost here."),
    ("5090", "Cost of Goods Sold", "Cost Of Goods Sold", "", "",
     "Zoho's inventory-module cost account. Leave it for items the inventory module sells.",
     "Plot and housing releases go to 5020 and 5025, not here."),

    ("6000", "Administrative and operating expenses", "Expense", "", "",
     "Header only. The lines below are the expense note.",
     "Do not post to the header. A cost that is still inventory does not belong here."),
    ("6010", "Salaries and Employee Wages", "Expense", "Administrative and operating expenses", "6010",
     "Staff salaries, wages and allowances. Workbook 6010. Not director remuneration.",
     "Director remuneration is 6015. A dividend is not a salary."),
    ("6015", "Director remuneration", "Expense", "Administrative and operating expenses", "",
     "Remuneration charged for the directors. Kept apart from staff costs so the related-party note can be tied.",
     "Do not post a dividend or a loan repayment here."),
    ("6020", "Rent Expense", "Expense", "Administrative and operating expenses", "6020",
     "Office and site rent. Workbook 6020.",
     "A deposit that will be recovered is a prepayment, not rent."),
    ("6025", "Repairs and Maintenance", "Expense", "Administrative and operating expenses", "6020",
     "Repairs and maintenance of the office, the site and the equipment. Workbook 6020 includes maintenance with rent. Kept as its own line so a repair is not lost in rent.",
     "A cost that extends the life of an asset is capital, not a repair."),
    ("6030", "Consultant Expense", "Expense", "Administrative and operating expenses", "6030",
     "Professional, engineering and legal fees. Workbook 6030.",
     "Audit and tax advisory fees are 6040. Selling commissions are 6035."),
    ("6035", "Selling commissions", "Expense", "Administrative and operating expenses", "",
     "Commissions on land and housing sales. Added because the expense note reports them on their own, and they are too large to bury in professional fees.",
     "Do not net a commission against revenue."),
    ("6040", "Audit and tax advisory fees", "Expense", "Administrative and operating expenses", "6040",
     "Audit and tax advisory fees. Workbook 6040.",
     "The tax itself is 8010, not this account."),
    ("6050", "Bank Fees and Charges", "Expense", "Administrative and operating expenses", "6050",
     "Bank charges, COT and guarantee fees. Workbook 6050.",
     "Interest on a facility is 7010, not a bank charge."),
    ("6055", "Credit Card Charges", "Expense", "Administrative and operating expenses", "",
     "Charges on a company card.",
     "The purchase on the card goes to the account for what was bought, not here."),
    ("6060", "Automobile Expense", "Expense", "Administrative and operating expenses", "6060",
     "Fuel, servicing and other running costs of vehicles. Workbook 6060.",
     "The purchase of a vehicle is 1520."),
    ("6061", "Travel Expense", "Expense", "Administrative and operating expenses", "",
     "Business travel other than vehicle running costs.",
     "Do not use this for a site transfer of materials."),
    ("6062", "Lodging", "Expense", "Administrative and operating expenses", "",
     "Lodging while on business travel. Zoho sample account.",
     "Staff housing that is part of remuneration is 6010, not lodging."),
    ("6070", "Depreciation Expense", "Expense", "Administrative and operating expenses", "6070",
     "Depreciation for the year. Workbook 6070. Credit 1590. Twenty per cent, straight line, full year in the year of purchase.",
     "Do not post the cost of an asset here."),
    ("6080", "Bad Debt", "Expense", "Administrative and operating expenses", "6080",
     "Impairment of receivables. Workbook 6080. Credit 1130.",
     "Do not write off a subscriber deposit. A deposit is a liability, not a receivable."),
    ("6085", "Customer refunds", "Expense", "Administrative and operating expenses", "",
     "Refunds that are not a reversal of Unearned Revenue or of a sale. Use this only where the receipt was already income and cannot be reversed there.",
     "A refund of a deposit reverses Unearned Revenue. It is not an expense dressed up as one."),
    ("6090", "Other Expenses", "Expense", "Administrative and operating expenses", "6090",
     "General office, security and utilities. Workbook 6090.",
     "Do not use this for a cost you can name. The named lines are above it."),
    ("6091", "Office Supplies", "Expense", "Administrative and operating expenses", "",
     "Stationery and other office supplies.",
     "Equipment that will last more than a year is 1510."),
    ("6092", "Advertising And Marketing", "Expense", "Administrative and operating expenses", "",
     "Advertising, marketing and selling events, other than a commission.",
     "A selling commission is 6035."),
    ("6093", "Telephone Expense", "Expense", "Administrative and operating expenses", "",
     "Telephone and mobile charges.",
     "A handset that is equipment is 1510."),
    ("6094", "IT and Internet Expenses", "Expense", "Administrative and operating expenses", "",
     "Internet and software subscriptions, including the Zoho Books subscription itself.",
     "The Zoho subscription is the company's cost, not a disbursement on behalf of a subscriber."),
    ("6095", "Janitorial Expense", "Expense", "Administrative and operating expenses", "",
     "Cleaning of the office and the site office.",
     "Site labour on a unit under construction is 5030."),
    ("6096", "Postage", "Expense", "Administrative and operating expenses", "",
     "Postage and courier.",
     "Do not use this for a shipping charge billed to a customer."),
    ("6097", "Printing and Stationery", "Expense", "Administrative and operating expenses", "",
     "Printing, including allocation letters and handover packs.",
     "A documentation fee charged to a subscriber is income, 4040, not a reduction of this expense."),
    ("6098", "Meals and Entertainment", "Expense", "Administrative and operating expenses", "",
     "Meals and entertainment.",
     "Do not use this for a staff allowance that is remuneration."),
    ("6088", "Purchase Discounts", "Expense", "Administrative and operating expenses", "",
     "Discount a vendor gives on a purchase. Zoho's type for this sample account is Expense, so it cannot sit under Cost of sales.",
     "A discount given to a subscriber is the Discount account, not this one."),
    ("9090", "Uncategorized", "Expense", "", "",
     "Zoho sample account. Do not use it.",
     "An unidentified payment is 1180. An unidentified receipt is 2045. Neither is an expense."),

    ("7000", "Finance costs", "Expense", "", "",
     "Header only.",
     "Do not post a bank charge here."),
    ("7010", "Finance costs - interest and similar charges", "Expense", "Finance costs", "7010",
     "Interest and similar charges on borrowings. Workbook 7010.",
     "COT and transfer charges are Bank Fees and Charges."),

    ("8000", "Income tax expense", "Expense", "", "",
     "Header only.",
     "Do not post to the header."),
    ("8010", "Income tax expense - current tax", "Expense", "Income tax expense", "8010",
     "Current tax for the year: companies income tax, tertiary education tax, the police trust fund levy and NASENI. Workbook 8010. Credit 2110 to 2140 for the levy each amount belongs to.",
     "WHT suffered is not tax expense. It is 1150."),
    ("8020", "Income tax expense - deferred tax", "Expense", "Income tax expense", "8020",
     "Deferred tax charge or credit. Workbook 8020. Use only if deferred tax is recognised.",
     "Do not use this for the current-year levy."),

    ("9000", "Other gains and losses", "Other Expense", "", "",
     "Header only.",
     "Do not post an operating expense here."),
    ("9010", "Exchange Gain or Loss", "Other Expense", "", "",
     "Exchange difference on a foreign-currency balance. Zoho sample account.",
     "Do not take an exchange difference to revenue or to a subscriber account."),
]


def accounts():
    out = []
    for code, name, typ, parent, workbook, post, do_not in ROWS:
        origin = "Existing" if name in SAMPLE_NAMES else "New"
        upload_parent = "" if name in LOCKED else parent
        out.append({
            "code": code,
            "name": name,
            "type": typ,
            "parent": parent,
            "upload_parent": upload_parent,
            "workbook": workbook,
            "origin": origin,
            "post": post,
            "do_not": do_not,
        })
    return out


def validate(rows):
    codes = [r["code"] for r in rows]
    names = [r["name"] for r in rows]
    if len(codes) != len(set(codes)):
        raise SystemExit(f"duplicate code: {[c for c in codes if codes.count(c) > 1]}")
    if len(names) != len(set(names)):
        raise SystemExit("duplicate name")
    if any(len(c) != 4 or not c.isdigit() for c in codes):
        raise SystemExit("a code is not 4 digits")
    missing = SAMPLE_NAMES - set(names)
    extra_existing = {r["name"] for r in rows if r["origin"] == "Existing"} - SAMPLE_NAMES
    if missing or extra_existing:
        raise SystemExit(f"sample mismatch missing={missing} extra={extra_existing}")
    by_name = {r["name"]: r for r in rows}
    seen = set()
    for r in rows:
        if r["upload_parent"]:
            parent = by_name.get(r["upload_parent"])
            if parent is None:
                raise SystemExit(f"{r['code']} parent missing from upload: {r['upload_parent']}")
            if parent["type"] != r["type"]:
                raise SystemExit(f"{r['code']} type {r['type']} != parent {parent['type']}")
            if r["upload_parent"] not in seen:
                raise SystemExit(f"{r['code']} parent {r['upload_parent']} is listed after the child")
        if r["parent"] and r["parent"] not in by_name:
            raise SystemExit(f"{r['code']} review parent missing: {r['parent']}")
        seen.add(r["name"])
    # Every workbook posting code from the controlled chart is either used as a
    # code or cited in the workbook column.
    workbook_codes = {
        "1010", "1020", "1030", "1040", "1050", "1110", "1120", "1130", "1140",
        "1150", "1160", "1210", "1220", "1230", "1510", "1520", "1530", "1590",
        "1610", "2010", "2020", "2030", "2040", "2110", "2120", "2130", "2140",
        "2150", "2160", "2510", "2610", "3010", "3020", "4010", "4020", "4030",
        "5010", "5020", "5030", "6010", "6020", "6030", "6040", "6050", "6060",
        "6070", "6080", "6090", "7010", "8010", "8020",
    }
    cited = set(codes) | {r["workbook"] for r in rows if r["workbook"]}
    absent = workbook_codes - cited
    if absent:
        raise SystemExit(f"workbook codes not represented: {sorted(absent)}")


HEADERS = [
    "Account ID",
    "Account Name",
    "Account Code",
    "Description",
    "Account Type",
    "Mileage Rate",
    "Mileage Unit",
    "IsMileage",
    "Account # ",
    "Account Status",
    "Currency",
    "Parent Account",
]


def description(r):
    text = r["post"]
    if r["do_not"]:
        text = f"{text} {r['do_not']}"
    return text


def write_csv(rows):
    with OUT_CSV.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)
        for r in rows:
            writer.writerow([
                "",
                r["name"],
                r["code"],
                description(r),
                r["type"],
                "0",
                "",
                "FALSE",
                r["code"],
                "Active",
                "NGN",
                r["upload_parent"],
            ])


NAVY = "1F4E79"
RED = "C00000"
INK = "1C2830"
WHITE = "FFFFFF"
PALE = "F4F7F8"
PARENT_FILL = "D6E3F0"
GOLD = "F4E8D4"
LINE = "D0D7E2"


def font(name, size, bold=False, color=INK, italic=False):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)


def apply_border(cell):
    thin = Border(
        left=Side(style="thin", color=LINE),
        right=Side(style="thin", color=LINE),
        top=Side(style="thin", color=LINE),
        bottom=Side(style="thin", color=LINE),
    )
    cell.border = thin


def write_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Zoho import"
    ws.sheet_properties.tabColor = NAVY
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:L{len(rows) + 1}"
    ws.row_dimensions[1].height = 22
    for col, header in enumerate(HEADERS, 1):
        cell = ws.cell(1, col, header)
        cell.font = font("Calibri", 10, True, WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        apply_border(cell)
    for i, r in enumerate(rows, 2):
        values = [
            "",
            r["name"],
            r["code"],
            description(r),
            r["type"],
            0,
            "",
            "FALSE",
            r["code"],
            "Active",
            "NGN",
            r["upload_parent"],
        ]
        is_parent = any(other["upload_parent"] == r["name"] for other in rows)
        if r["origin"] == "Existing":
            colour = RED
            fill = "FDF2F2"
        elif is_parent:
            colour = NAVY
            fill = PARENT_FILL
        else:
            colour = INK
            fill = WHITE if i % 2 else PALE
        for col, value in enumerate(values, 1):
            cell = ws.cell(i, col, value)
            cell.font = font("Calibri", 9, bold=is_parent or col in (2, 3), color=colour)
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.alignment = Alignment(vertical="center", wrap_text=(col == 4))
            apply_border(cell)
            if col in (3, 9):
                cell.number_format = "@"
        lines = max(2, (len(description(r)) // 85) + 1)
        ws.row_dimensions[i].height = min(72, 15 * lines + 6)
    widths = [14, 52, 14, 78, 28, 14, 14, 12, 12, 14, 12, 42]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.auto_filter.ref = f"A1:L{len(rows) + 1}"
    ws.oddHeader.left.text = "Baay Projects Limited — Zoho Books chart of accounts"
    ws.oddFooter.left.text = "Existing accounts are in red. Import this sheet, or the CSV, with Overwrite."
    ws.oddFooter.right.text = "Page &P of &N"
    ws.page_setup.horizontalCentered = True
    ws.sheet_view.zoomScale = 110
    ws.print_title_rows = "1:1"
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.7
    ws.page_margins.bottom = 0.6

    review = wb.create_sheet("Review")
    review.sheet_properties.tabColor = "8C4A3A"
    review.sheet_view.showGridLines = False
    review.freeze_panes = "A2"
    review_headers = [
        "Code",
        "Account name",
        "In the Zoho sample",
        "Account type",
        "Parent in the upload file",
        "Intended reporting group",
        "Workbook code",
        "What to post",
        "What not to post",
    ]
    review.row_dimensions[1].height = 22
    for col, header in enumerate(review_headers, 1):
        cell = review.cell(1, col, header)
        cell.font = font("Calibri", 10, True, WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        apply_border(cell)
    for i, r in enumerate(rows, 2):
        group = r["parent"] or ("Locked Zoho account — leave where Zoho holds it" if r["name"] in LOCKED else "Header")
        values = [
            r["code"],
            r["name"],
            "Existing — marked red" if r["origin"] == "Existing" else "New",
            r["type"],
            r["upload_parent"],
            group,
            r["workbook"],
            r["post"],
            r["do_not"],
        ]
        colour = RED if r["origin"] == "Existing" else INK
        fill = "FDF2F2" if r["origin"] == "Existing" else (WHITE if i % 2 else PALE)
        for col, value in enumerate(values, 1):
            cell = review.cell(i, col, value)
            cell.font = font("Calibri", 9, bold=(col in (1, 2) or r["origin"] == "Existing"), color=colour)
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.alignment = Alignment(vertical="center", wrap_text=col in (8, 9))
            apply_border(cell)
            if col in (1, 7):
                cell.number_format = "@"
        lines = max(2, (len(r["post"]) // 70) + 1)
        review.row_dimensions[i].height = min(64, 14 * lines + 6)
    review.auto_filter.ref = f"A1:I{len(rows) + 1}"
    for i, width in enumerate([10, 52, 24, 28, 42, 46, 16, 62, 62], 1):
        review.column_dimensions[get_column_letter(i)].width = width
    review.auto_filter.ref = f"A1:I{len(rows) + 1}"
    review.page_setup.orientation = "landscape"
    review.page_setup.fitToPage = True
    review.page_setup.fitToWidth = 1
    review.page_setup.fitToHeight = 0
    review.page_setup.paperSize = review.PAPERSIZE_A3
    review.print_title_rows = "1:1"
    review.page_margins.left = 0.4
    review.page_margins.right = 0.4
    review.oddFooter.right.text = "Page &P of &N"

    notes = wb.create_sheet("How to import", 0)
    notes.sheet_properties.tabColor = "B08968"
    notes.sheet_view.showGridLines = False
    notes.page_setup.orientation = "portrait"
    notes.page_setup.fitToPage = True
    notes.page_setup.fitToWidth = 1
    notes.page_setup.fitToHeight = 0
    notes.page_setup.paperSize = notes.PAPERSIZE_A4
    notes.column_dimensions["A"].width = 3
    notes.column_dimensions["B"].width = 108
    notes.row_dimensions[1].height = 10
    notes.merge_cells("B2:B2")
    title = notes["B2"]
    title.value = "Baay Projects Limited"
    title.font = font("Calibri", 20, True, NAVY)
    notes.row_dimensions[2].height = 26
    sub = notes["B3"]
    sub.value = "Chart of accounts for Zoho Books"
    sub.font = font("Calibri", 14, True, "1F6F78")
    notes.row_dimensions[3].height = 20
    kicker = notes["B4"]
    kicker.value = "Four-digit codes. Existing Zoho accounts are in red on the import sheet. New accounts are in black."
    kicker.font = font("Calibri", 11, False, "5E6A72")
    notes.row_dimensions[4].height = 18

    blocks = [
        ("What this file is",
         "The import sheet, and the CSV of the same name, are the file to load into Zoho Books. The Review sheet is for the person posting. It is not imported. Colour cannot be stored in a CSV, which is why the red mark is in this workbook."),
        ("Where the detail comes from",
         "The level of detail is the controlled chart in the master workbook for 2023 to 2025: one account for each bank, receivables and revenue split by stream, inventory in three buckets, tax split by levy, and expenses split by nature. Estates are not accounts. They are Zoho projects. A further account was added only where the sales register already separates the line, or where a balance cannot be posted without one: housing, ancillary charges, co-ownership, unidentified receipts and payments, related parties, and the unpaid amount on issued shares."),
        ("How to import",
         "In Zoho Books go to Accountant, Chart of Accounts, and Import Chart of Accounts. Upload the CSV. Duplicate handling must be Overwrite, so the four-digit codes land on the accounts already in the organisation. Zoho matches those accounts on the name. Character encoding is UTF-8. Map Account Name, Account Code, Description, Account Type, Parent Account, Currency and Account Status. Do not map Account ID. The sample export's IDs were damaged by Excel and are not in this file. Do not import this instruction sheet."),
        ("Parents and children",
         "A child is the same account type as its parent. That is a Zoho rule, and this file already obeys it. Parents are listed before their children. A header account is not a posting account. Once it has children, Zoho will not accept a transaction on it."),
        ("Accounts Zoho will not re-parent",
         "Accounts Receivable, Accounts Payable, Retained Earnings, Opening Balance Offset, Unearned Revenue, VAT Payable, Undeposited Funds, Inventory Asset, Cost of Goods Sold, Sales, Discount, Shipping Charge, Other Charges, Late Fee Income, Opening Balance Adjustments, Dimension Adjustments and Exchange Gain or Loss are coded, and marked red, but they have no parent in the upload file. Zoho locks these, and a parent on the import is the usual reason a row fails. Their intended group is on the Review sheet. After the import, move one only if the edit screen allows it. Do not force it. Do not post estate transactions to Sales, Cost of Goods Sold or Inventory Asset."),
        ("How a naira is treated",
         "A receipt from a subscriber is Unearned Revenue until the plot is allocated or the house is handed over. Co-ownership is a liability, never revenue. A movement between Providus, First Bank, Sterling and GTBank is a transfer. A receipt or payment whose nature is not known waits in 2045 or 1180. It is not posted to profit to close a month. Residential land and housing are not charged to VAT. Depreciation is 20 per cent a year, straight line, with a full year in the year of purchase."),
        ("Opening balances",
         "Open from the signed statement of financial position at 31 December 2025. This chart does not contain those figures. Opening Balance Offset and Opening Balance Adjustments should both be nil when the opening trial balance ties."),
        ("Count",
         f"{len(rows)} accounts. {sum(r['origin']=='Existing' for r in rows)} are already in the Zoho sample and are shown in red. {sum(r['origin']=='New' for r in rows)} are new."),
    ]
    row = 6
    for heading, body in blocks:
        notes.cell(row, 2, heading).font = font("Calibri", 12, True, NAVY)
        notes.row_dimensions[row].height = 18
        row += 1
        cell = notes.cell(row, 2, body)
        cell.font = font("Calibri", 10.5, False, INK)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        lines = max(2, (len(body) // 105) + 1)
        notes.row_dimensions[row].height = 16 * lines + 8
        row += 2
    notes.row_dimensions[row].height = 8
    legend = notes.cell(row + 1, 2, "Red text on the Zoho import sheet and the Review sheet means the account was already in the Zoho sample. Navy bold means a header. Black means a new posting account.")
    legend.font = font("Calibri", 10, True, RED)
    legend.alignment = Alignment(wrap_text=True)
    notes.row_dimensions[row + 1].height = 32
    notes.print_title_rows = "1:3"
    notes.page_margins.left = 0.6
    notes.page_margins.right = 0.6
    notes.page_margins.top = 0.6
    notes.page_margins.bottom = 0.6
    notes.oddFooter.left.text = "Baay Projects Limited — Zoho Books chart of accounts"
    notes.oddFooter.right.text = "Confidential"

    # The import sheet must be first if someone imports the workbook by mistake
    # after reading the notes. Keep notes first for the person opening the file,
    # and tell them to use the CSV. That is the safer instruction.
    wb.active = notes
    wb.properties.title = "Baay Projects Limited — Zoho Books chart of accounts"
    wb.properties.creator = "CloudERP Accounting & Business Solutions"
    wb.properties.subject = "Chart of accounts for upload to Zoho Books. Existing accounts marked in red."
    wb.save(OUT_XLSX)


def main():
    rows = accounts()
    validate(rows)
    write_csv(rows)
    write_xlsx(rows)
    print(f"accounts {len(rows)} existing {sum(r['origin']=='Existing' for r in rows)} new {sum(r['origin']=='New' for r in rows)}")
    print("wrote", OUT_CSV)
    print("wrote", OUT_XLSX)


if __name__ == "__main__":
    main()
