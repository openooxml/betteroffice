---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
"@betteroffice/rust-crates": minor
---

The accessible page text reads a line break inside a paragraph as a new line and a tab as a tab, so the words on either side no longer run together for screen readers. Rust `DocAttrs` literals must set `separator_before` and `separator_after`.
