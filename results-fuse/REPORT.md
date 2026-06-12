# xlsx corpus benchmark — results

Universe: 10,544 workbooks — the 10,702 unique files minus 158 that Excel itself could not process; every library is measured on exactly this set, and a library failure inside it counts against that library rather than shrinking its denominator. See README.md for methodology and definitions.

Recalculation is judged against real Excel: Microsoft Excel recomputed every workbook (harness/excel_truth.py) and each engine's results are compared to Excel's by one shared comparator (harness/compare_truth.py). The two recalculation columns answer different questions: how many *workbooks* come out perfect (all-or-nothing per file), and how many *individual formulas* match (overall accuracy; a few big failing workbooks can move this a lot without moving the first).

| library | workbooks | opens without error | survives open→save→reopen | workbooks recalculated 100% Excel-identical | formula cells matching Excel |
|---|---|---|---|---|---|
| closedxml | 10,544 | 99.7% | 98.6% | 82.1% of 3,577 | 69.0% of 5,671,240 |
| epplus | 10,544 | 98.6% | 98.6% | 62.3% of 3,577 | 52.2% of 5,671,240 |
| libreoffice | 10,544 | 100.0% | 94.8% | 91.9% of 3,577 | 98.9% of 5,671,240 |
| openpyxl | 10,544 | 99.9% | 99.9% | N/A (no calculation engine) | N/A (no calculation engine) |
| witan | 10,544 | 100.0% | 100.0% | 95.3% of 3,577 | 99.7% of 5,671,240 |

## Top load-error signatures

**closedxml**
- 5 × `ArgumentException: Unable to determine the format of the image.`
- 4 × `ArgumentException: The picture ID '0' already exists.`
- 3 × `ArgumentException: sheetName must not be null or whitespace`
- 3 × `ArgumentOutOfRangeException: Table  was not found. (Parameter 'name')`
- 2 × `PartStructureException: There is a problem with element structure in XML, the number of elements found is not what was e`

**epplus**
- 135 × `NotSupportedException: This property or method is not supported for a Chartsheet`
- 9 × `timeout`
- 2 × `FormatException: The input string '' was not in a correct format.`
- 2 × `ArgumentOutOfRangeException: Specified argument was out of the range of valid values. (Parameter 'Start cell Address mus`
- 1 × `XmlException: 'o:relid' is a duplicate attribute name. Line 31, position 26.`

**libreoffice**
- 1 × `conversion produced no output`

**openpyxl**
- 2 × `TypeError: expected <class 'float'>`
- 1 × `ValueError: Unable to read workbook: could not read strings from corpus-fuse/1189c182-f659-4493-989c-76be4b227aa7.xlsx.`
- 1 × `ValueError: Unable to read workbook: could not read strings from corpus-fuse/27ba685d-a4ae-4ae7-a0c9-9d09d586d745.xlsx.`
- 1 × `ValueError: Unable to read workbook: could not read strings from corpus-fuse/427f5d2b-d5d7-48d7-8dcc-9ef912934f5c.xlsx.`
- 1 × `ValueError: Unable to read workbook: could not read strings from corpus-fuse/8c0368c2-8b06-4c12-818d-28a18be3381b.xlsx.`

