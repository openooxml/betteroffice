---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

Saving after a style edit no longer copies conditional-format number formats into the cell number formats, where one could redefine a built-in format, such as a date, for untouched cells. Conditional-format styles save unchanged.
