---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Line breaks stored as CR LF or a lone CR read as one line feed, as in Excel, so table references to multi-line column names calculate instead of returning #REF!. Saves keep carriage returns in cell text, formulas and hyperlink tooltips.
