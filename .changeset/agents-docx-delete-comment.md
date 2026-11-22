---
'@betteroffice/agents': patch
'@betteroffice/docx': patch
'@betteroffice/rust-crates': minor
---

`docx_edit` adds `delete_comment`, which removes a comment or a root comment's whole thread, including its range markers, reference and companion entries. Rust `S13ParagraphIds` literals must set `removed_comment_ids`.
