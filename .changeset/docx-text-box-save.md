---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Text boxes now keep their text when a document is saved. Shape text bodies read from `wps:txbx` were dropped on the way out — including the `mc:AlternateContent` form Word writes for figure callouts — so every text box came back empty.
