---
"@betteroffice/python-xlsx": patch
---

`render_png` now draws XLSX columns at Excel's on-screen widths, using the sheet's `defaultColWidth` or `baseColWidth` for columns without their own width, so rendered images can change.
