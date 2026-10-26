---
"@betteroffice/xlsx": patch
---

Sheets styled on more than 65,536 rows now open, edit and save instead of failing with "worksheet row style count exceeded cap". Live collaboration on a sheet with more than about 333,000 styled rows is refused for now.
