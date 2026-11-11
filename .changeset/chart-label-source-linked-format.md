---
"@betteroffice/xlsx": patch
"@betteroffice/pptx": patch
"@betteroffice/docx": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/python-pptx": patch
"@betteroffice/python-docx": patch
"@betteroffice/rust-crates": patch
---

Chart data labels and value axes linked to their source cells use the format the values were saved with, so a `"$"#,##0` label reads $3,750. Pie percentage labels use their own number format.
