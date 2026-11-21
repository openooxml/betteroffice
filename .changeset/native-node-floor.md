---
'@betteroffice/docx-native': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/pptx-native': minor
'@betteroffice/agents': minor
---

The native packages now export the full `@betteroffice/docx`, `@betteroffice/xlsx` and `@betteroffice/pptx` APIs as ES modules, and `require()` of them needs Node 20.19 or 22.12: `engines` now says so. `@betteroffice/agents`, which depends on them, requires Node 22.12.
