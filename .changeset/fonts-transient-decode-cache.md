---
"@betteroffice/fonts": patch
"@betteroffice/xlsx-react": patch
---

The spreadsheet editor no longer keeps decoded copies of bundled fonts after the browser has registered them, which cuts page memory by about 2.5 to 5 MB.
