---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Spreadsheet cells load bundled fallback fonts for scripts such as Arabic, Hebrew, Thai and Devanagari, so text stays readable when the device lacks those fonts. PNG exports from the editor use the same fonts.
