---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Deleting a comment in the editor and saving keeps the source bytes of the other paragraphs when the save can patch the parts, which it cannot when the same save also carries a host-side change such as page margins. Comment and range ids match by value.
