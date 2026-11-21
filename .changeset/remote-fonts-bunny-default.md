---
"@betteroffice/fonts": minor
"@betteroffice/docx": minor
"@betteroffice/xlsx-react": minor
"@betteroffice/python-docx": minor
"@betteroffice/python-xlsx": minor
"@betteroffice/python-pptx": minor
"@betteroffice/rust-crates": patch
---

DOCX no longer defaults to Google Fonts. Unbundled families share one substitution table. XLSX fetches only substitutes' visible-text character-range files from Bunny Fonts via `remoteFonts` or `setRemoteFontProvider` ('bunny', 'google', custom URL, 'off'). Python adds `font_substitute`. In DOCX, `setRemoteFontProvider` and `setGoogleFontsEnabled` have no effect.
