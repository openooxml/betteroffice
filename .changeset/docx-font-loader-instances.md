---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Editors sharing a page keep fonts apart: `onFontsLoaded` and `onError` no longer hear other editors' font loads, and different embedded fonts under one name no longer override each other and are released on unmount. Adds `createFontLoadScope`, `registerDocumentFaces` and `loadEmbeddedFontFamilies`.
