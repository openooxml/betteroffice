---
'@betteroffice/docx': minor
'@betteroffice/pptx': minor
'@betteroffice/xlsx': minor
'@betteroffice/vsdx': minor
'@betteroffice/docx-native': patch
'@betteroffice/pptx-native': patch
'@betteroffice/xlsx-native': patch
---

Adds opt-in `decodeFontFace` for packed fonts. Format package roots return a promise of `font` and optional `payload` ArrayBuffers; native bindings return byte arrays synchronously. Invalid options and decode errors reject or throw without retrying another decoder.
