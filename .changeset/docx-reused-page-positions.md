---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Typing early in a long document no longer resolves the text of every later page again, because `layout_document_incremental` now moves the run positions of retained pages by the edit's offset and takes equal fingerprints to mean a block is unchanged apart from its document positions.
