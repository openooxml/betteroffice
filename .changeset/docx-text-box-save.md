---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Text boxes now keep their text when a document is saved, in the editor as well as the crate — the editor's save projection rebuilt every shape as an empty rectangle. Shape text bodies read from `wps:txbx` were dropped on the way out — including the `mc:AlternateContent` form Word writes for figure callouts, a table nested in a box, and a vertical (`vert="eaVert"`) writing direction — so every text box came back empty.
