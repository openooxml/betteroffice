---
"@betteroffice/xlsx": patch
---

A workbook too large for the spreadsheet engine's memory now fails the session with an `out-of-memory` failure that says the workbook is too large, instead of a bare `trap` reading "unreachable", whether it runs out while opening or later.
