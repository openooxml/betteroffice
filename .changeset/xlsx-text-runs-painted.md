---
"@betteroffice/xlsx": minor
"@betteroffice/rust-crates": minor
"@betteroffice/xlsx-native": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/agents": patch
---

XLSX text draw commands gain optional `runs`, spans of the text in their own font, size, style and colour; the canvas, both rasters and the native viewer paint them on single-line text. Nothing emits runs yet, so sheets look the same.
