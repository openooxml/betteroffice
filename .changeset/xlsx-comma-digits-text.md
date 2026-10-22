---
"@betteroffice/xlsx": patch
---

Digits separated only by commas, such as `1,2,3` or `12,31,2024`, typed into a cell now stay text, as in Excel, instead of becoming a date. Month-name dates with commas still parse.
