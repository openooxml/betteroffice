---
'@betteroffice/docx': patch
'@betteroffice/pptx': patch
'@betteroffice/xlsx': patch
'@betteroffice/vsdx': patch
'@betteroffice/docx-native': patch
'@betteroffice/pptx-native': patch
'@betteroffice/xlsx-native': patch
---

Adds opt-in `decodeFontFace` engine calls for zstd and woff2t font pack faces, returning font bytes and an optional dictionary payload. Non-zstd bytes pass through.
