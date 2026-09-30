---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": minor
---

Embedded images stay compressed in the opened file until a page shows them. `createCanvasImageResolver` accepts a `media` source, and the new `mediaTokens` option (`DocxEditor`, `YrsOpeningOptions`), off by default, keeps images out of the shared document state.
