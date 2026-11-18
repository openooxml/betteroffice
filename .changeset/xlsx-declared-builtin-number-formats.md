---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/xlsx-react': patch
---

A number format a workbook declares under a built-in id is shown as declared, in cell text, the structured export, `CELL("format")` and typed input: an accounting format with a minus sign or an East Asian date no longer shows the built-in reading. General keeps its reading whatever the workbook declares for it. Shared states from earlier releases keep the built-in display until the workbook is opened again. Collaborating peers must update together: older peers refuse the new schema.
