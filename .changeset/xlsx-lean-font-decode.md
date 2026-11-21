---
"@betteroffice/xlsx": patch
---

Spreadsheet editor sessions no longer download and start a second, full copy of the spreadsheet engine when they decode their first font, saving about 2.9 MB of compressed download per open.
