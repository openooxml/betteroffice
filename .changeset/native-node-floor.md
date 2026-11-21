---
'@betteroffice/docx-native': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/pptx-native': minor
---

The native packages now export the full `@betteroffice/docx`, `@betteroffice/xlsx` and `@betteroffice/pptx` APIs as ES modules. `require()` of them needs Node 20.19 or 22.12, so `engines` now requires `^20.19.0 || >=22.12.0`.
