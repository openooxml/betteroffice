---
"@betteroffice/xlsx": patch
---

`TEXT` now formats text that reads as a number or date, as Excel does, so `TEXT` over a cell holding "November 2025" with a date format gives the formatted date.
