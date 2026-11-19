---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

A document reopened with the same `clientId`, such as after a page reload, now gives new comments, tracked changes and paragraphs ids of their own, instead of reusing ids the document already holds.
