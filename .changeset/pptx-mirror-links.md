---
"@betteroffice/pptx": patch
"@betteroffice/pptx-native": patch
"@betteroffice/pptx-react": patch
"@betteroffice/rust-crates": minor
---

Adds paragraph `links` to slide display lists. Screen reader slide mirrors expose safe external links and slide navigation as keyboard accessible links by default. Rust `TextParagraph` literals must set `links`.
