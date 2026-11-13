---
"@betteroffice/xlsx": patch
"@betteroffice/pptx": patch
"@betteroffice/docx": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/python-pptx": patch
"@betteroffice/python-docx": patch
"@betteroffice/rust-crates": minor
---

Chart data labels and value axes linked to their source cells use the format the values were saved with, so a `"$"#,##0` label reads $3,750; without one they keep their own code. Pie percentages read a label's own format, and a plain percentage when the label is linked to its source.
