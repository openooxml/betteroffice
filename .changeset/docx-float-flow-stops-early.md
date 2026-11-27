---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Editing next to text-wrapped floating objects in DOCX lays out faster, because layout after the edit stops once the following text is unchanged. Section grid changes beside such objects are respected.
