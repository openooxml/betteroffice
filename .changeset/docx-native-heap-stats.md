---
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
'@betteroffice/docx-react': patch
---

The native DOCX engine can report its editing heap per thread. Counting is off by default; `startHeapCounting()` from `@betteroffice/docx/yrs` turns it on. Pressure events on native need a host memory budget.
