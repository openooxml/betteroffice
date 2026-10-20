---
"@betteroffice/pptx": patch
"@betteroffice/xlsx": patch
"@betteroffice/docx": patch
---

Charts draw a quoted or escaped `%` in a number format (`0.0"%"`, `0.0\%`, `0.0[$%]`) as literal text instead of scaling the value by 100 and adding a second `%`.
