# xlsx corpus benchmark — results

Universe: 5,426 workbooks — the 5,455 unique files minus 29 that Excel itself could not process; every library is measured on exactly this set, and a library failure inside it counts against that library rather than shrinking its denominator. See README.md for methodology and definitions.

Recalculation is judged against real Excel: Microsoft Excel recomputed every workbook (harness/excel_truth.py) and each engine's results are compared to Excel's by one shared comparator (harness/compare_truth.py). The two recalculation columns answer different questions: how many *workbooks* come out perfect (all-or-nothing per file), and how many *individual formulas* match (overall accuracy; a few big failing workbooks can move this a lot without moving the first).

| library | workbooks | opens without error | survives open→save→reopen | workbooks recalculated 100% Excel-identical | formula cells matching Excel |
|---|---|---|---|---|---|
| closedxml | 5,426 | 99.5% | 98.7% | 63.8% of 2,970 | 39.4% of 1,168,856 |
| epplus | 5,426 | 99.8% | 99.5% | 74.6% of 2,970 | 68.0% of 1,168,856 |
| libreoffice | 5,426 | 100.0% | 97.3% | 90.4% of 2,970 | 96.7% of 1,168,856 |
| openpyxl | 5,426 | 100.0% | 99.8% | N/A (no calculation engine) | N/A (no calculation engine) |
| witan | 5,426 | 100.0% | 100.0% | 94.8% of 2,970 | 99.8% of 1,168,856 |

## Top load-error signatures

**closedxml**
- 18 × `PartStructureException: There is a problem with element structure in XML, the number of elements found is not what was e`
- 3 × `ParsingException: Unable to determine token for ''D:\文档\verify_table\4.18\31202\Balance Gastos 2024.xlsm'!tblProveedores`
- 3 × `ParsingException: Unable to determine token for ''F:\zhipuwork\excel公式标注\excel任务二质检\418_谢蓉_24\31202\Balance Gastos 2024.`
- 3 × `timeout`

**epplus**
- 11 × `timeout`

