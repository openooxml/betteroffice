---
"@betteroffice/xlsx": patch
---

Inserting or deleting rows or columns no longer blocks saving a workbook whose chart caches hold content the engine does not model, such as per-point number formats. Such a cache is kept as authored while its cells read the same; otherwise the refusal names the chart and series.
