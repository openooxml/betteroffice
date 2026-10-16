---
"@betteroffice/xlsx": patch
---

Saving an edited workbook that has no formulas now keeps its calculation settings as they were, so spreadsheet apps no longer recalculate it in full on open. Workbooks with formulas still request a full recalculation.
