---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Deleting a comment in the editor and saving now changes only the paragraphs that held it and its replies, even where the document already repeats a bookmark id, unless such a paragraph also opens another comment that ends later. Other paragraphs keep their source bytes.
