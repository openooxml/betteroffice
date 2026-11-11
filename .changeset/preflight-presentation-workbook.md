---
"@betteroffice/pptx": minor
"@betteroffice/xlsx": minor
"@betteroffice/pptx-native": minor
"@betteroffice/xlsx-native": minor
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
---

Adds experimental `fontPreflight` and `preloadPreflightWasm` for the active slide and the opening sheet, without opening the editing engine. DOCX `fontPreflight` answers now follow the engine for notes, the main part, numbering, theme script fonts, style defaults and table columns.
