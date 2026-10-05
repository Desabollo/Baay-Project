# BAAY Projects Limited — 12-Point Audit Integrity Re-verification

**Scope:** 36-sheet master workbook after full ingestion of all eight accountant-workbook tabs.
**Result:** PASS — all 12 checks passed.  
**Formula inventory:** 309 formulas in the integrated master; 2,576 formulas parsed in the accountant source.  
**Plug control:** No unexplained balancing journal or unsupported sales-to-revenue posting was introduced. The sales-to-bank allocation difference is disclosed as a derived reconciliation control, not posted into the general ledger.

| Check | Verification | Expected | Actual | Difference | Status |
|---|---|---:|---:|---:|---|
| CHK-01 | Sales ledger total tie | 2,584,814,000.00 | 2,584,814,000.00 | 0.00 | PASSED |
| CHK-02 | Annual sales roll-up | 2,584,814,000.00 | 2,584,814,000.00 | 0.00 | PASSED |
| CHK-03 | Exclusions completeness | 238,933,750.00 | 238,933,750.00 | 0.00 | PASSED |
| CHK-04 | Turnover source-sheet tie | 0.00 | 0.00 | 0.00 | PASSED |
| CHK-05 | Sales transaction population | 859.00 | 859.00 | 0.00 | PASSED |
| CHK-06 | Subscription population | 214.00 | 214.00 | 0.00 | PASSED |
| CHK-07 | Customer master population | 236.00 | 236.00 | 0.00 | PASSED |
| CHK-08 | Master workbook structure | 36.00 | 36.00 | 0.00 | PASSED |
| CHK-09 | Adjusted trial balance | 0.00 | 0.00 | 0.00 | PASSED |
| CHK-10 | Statement of financial position | 0.00 | 0.00 | 0.00 | PASSED |
| CHK-11 | Formula integrity | 0.00 | 0.00 | 0.00 | PASSED |
| CHK-12 | Receipt-to-bank control bridge | 0.00 | 0.00 | 0.00 | PASSED |

## Reconciliation interpretation

The accountant workbook describes its amounts as customer cash receipts rather than IFRS 15 revenue. Accordingly, all source populations are embedded in the master workbook, while recognized revenue remains controlled by performance-obligation evidence. The separate reconciliation workbook provides the annual bridge, complete sales ledger, exclusions, secondary-record matching population, turnover reconciliation and bank-account coverage schedule.
