---
"@betteroffice/xlsx": patch
---

Workbooks with SUMPRODUCT over long, mostly empty ranges now recalculate to Excel's results instead of showing `#NUM!`.
