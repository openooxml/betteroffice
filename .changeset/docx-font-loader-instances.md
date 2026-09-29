---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

By default, editors sharing a page keep fonts apart: `onFontsLoaded` and `onError` ignore other editors' font loads, and same-named embedded fonts no longer override each other and are released on unmount. Direct callers opt in via `createFontLoadScope`, `registerDocumentFaces` and `loadEmbeddedFontFamilies`.
