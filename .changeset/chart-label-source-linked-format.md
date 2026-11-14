---
"@betteroffice/xlsx": patch
"@betteroffice/pptx": patch
"@betteroffice/docx": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/python-pptx": patch
"@betteroffice/python-docx": patch
"@betteroffice/rust-crates": minor
---

Chart data labels and value axes linked to their source cells use the cached format of the values, so a `"$"#,##0` series labels $3,750; without a cached format they keep their own code. Linked pie percentages show a plain percentage.
