---
'@betteroffice/xlsx-react': minor
---

`XlsxEditor` with `readOnly` and `experimentalWorkerOpen` opens the workbook in a worker, which builds the frames the page paints, keeping the page responsive while large workbooks open. Off by default.
