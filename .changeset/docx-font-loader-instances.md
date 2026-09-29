---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Editors sharing a page keep fonts apart: `onFontsLoaded` and `onError` hear only their own editor's loads, and different embedded fonts under one name no longer override each other and are released on unmount. Adds `createFontLoadScope`, `registerDocumentFaces` and `loadEmbeddedFontFamilies`.
