---
"@betteroffice/docx": patch
---

Bundled fonts configured via `configureDefaultFonts` load from the bundle instead of Google Fonts, matched by Word name or bundled name (e.g. `Gelasio`); a font with no bundled CSS match still gets its own bundled faces (e.g. Comic Sans MS).
