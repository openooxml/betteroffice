---
'@betteroffice/pptx': patch
'@betteroffice/pptx-react': patch
'@betteroffice/rust-crates': minor
---

Adds paragraph `languages` to slide display lists. Screen readers use authored text languages by default, including mixed-language paragraphs and titles, while paragraphs with one language remain a single text node. Rust `TextParagraph` literals must set `languages`.
