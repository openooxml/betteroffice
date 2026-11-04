---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/docx-native': patch
---

Editing text inside a block content control, such as a form field, no longer rewrites the rest of the document on save: tracked changes, proofing marks, drawings and fields outside the edited paragraph are kept as they were.
