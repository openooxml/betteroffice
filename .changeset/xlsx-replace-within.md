---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-react': patch
'@betteroffice/rust-crates': minor
'@betteroffice/python-xlsx': minor
---

`replaceText` takes `within: 'workbook'`, which skips hidden, very hidden and non-worksheet sheets; an explicit chartsheet target is refused. Replace stats add `refused`, counting every refused cell, including any past the 10,000 listed.
