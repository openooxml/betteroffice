---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

The spreadsheet editor uses bundled fallbacks for missing glyphs in scripts such as Hebrew, Thai and Devanagari, with immediate retries after an editor closes mid-load. PNG exports also load script fallbacks for characters introduced by number formats.
