---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

When a DOCX is first loaded for editing, equal attribute keys and values share one allocation. Memory retained after open drops by 1 to 6% on large real documents (10% on a text-heavy synthetic 5 MiB file). Peak memory and saved bytes are unchanged.
