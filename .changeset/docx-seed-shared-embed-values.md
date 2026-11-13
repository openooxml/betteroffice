---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

When a DOCX is first loaded for editing, equal embed payload keys and values, such as the formatting carried by paragraph marks, share one allocation. Memory retained after open drops, and saved bytes are unchanged.
