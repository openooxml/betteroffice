---
"@betteroffice/pptx": patch
"@betteroffice/pptx-react": patch
"@betteroffice/rust-crates": minor
---

Plugin reads include `headerFooter` with visible slide footer, slide number and date text, plus stored settings in `layout.headerFooter`. These facts are included by default and optional for older engines. Rust `SlideSnapshot` and `SlideLayoutSnapshot` literals must set `header_footer`.
