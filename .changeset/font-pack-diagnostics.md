---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
---

`fontPackDiagnostics()` names installed font add-ons that cannot be used (another release, missing files, unreadable manifest). Native packages also warn once per pack with code `BETTEROFFICE_FONT_PACK`. Exports keep substituting those faces.
