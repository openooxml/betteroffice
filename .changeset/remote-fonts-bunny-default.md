---
"@betteroffice/fonts": minor
"@betteroffice/docx": minor
"@betteroffice/xlsx-react": minor
"@betteroffice/python-docx": minor
"@betteroffice/python-xlsx": minor
"@betteroffice/python-pptx": minor
"@betteroffice/rust-crates": patch
---

DOCX no longer loads Google Fonts by default. Unbundled families use one shared substitution table; XLSX fetches their look-alike from Bunny Fonts, configurable with `remoteFonts` or `setRemoteFontProvider` ('bunny', 'google', a custom URL, 'off'). Python adds `font_substitute`.
