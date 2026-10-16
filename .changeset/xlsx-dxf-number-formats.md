---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

Saving style edits preserves conditional-format styles without copying their number formats into cell formats. Workbooks with large conditional-format collections can open without a style-limit error, while saved collaborative state from earlier releases remains compatible.
