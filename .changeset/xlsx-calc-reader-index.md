---
"@betteroffice/xlsx": patch
---

Recalculation after an edit finds the formulas that read a cell through an index instead of scanning every dependency on the sheet, so a long chain such as a running total over 100,000 rows recalculates in time proportional to its length.
