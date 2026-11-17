---
"@betteroffice/docx": patch
"@betteroffice/fonts": patch
"@betteroffice/rust-crates": patch
---

DOCX text that no loaded font covers falls back to bundled BetterOffice Base Latin faces, loaded on demand from `@betteroffice/fonts`. Native raster builds embed the faces. Private-use bullet aliases apply only to Symbol and Wingdings families.
