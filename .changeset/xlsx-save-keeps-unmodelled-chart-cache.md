---
"@betteroffice/xlsx": patch
---

Inserting or deleting rows or columns no longer blocks saving workbooks whose chart caches hold unmodelled content, such as per-point number formats. The cache is kept while its cells read the same; otherwise the refusal names the chart and series.
