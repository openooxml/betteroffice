---
"@betteroffice/rust-crates": patch
---

Make the first keystroke after opening a document as incremental as later ones. The region reuse walk now compares section breaks without the margins a tall header widens them to, so the first edit converges after a page or two instead of repaginating the rest of a multi-section document, and the edit path keeps the frame extras the host wrote when they carry the same headers and footers, so it builds on the host's frame instead of rebuilding every page's display input.
