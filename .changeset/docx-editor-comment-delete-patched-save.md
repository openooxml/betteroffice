---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Deleting a comment inside one paragraph and saving in the editor keeps the other paragraphs' source bytes, unless the save also carries a host-side change such as page margins. A comment spanning paragraphs still rewrites the body. Comment and bookmark ids compare by value.
