---
'@betteroffice/docx': minor
'@betteroffice/docx-react': patch
---

Drag selection in the worker editor is cheaper per step: the worker keeps parsed pages, re-rasters no unchanged page and answers once without whole-frame geometry. `createDisplayListQueries` takes an optional `storePages` cap on the pages it keeps parsed.
