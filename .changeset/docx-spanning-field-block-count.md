---
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
'@betteroffice/rust-crates': minor
---

A field spanning paragraphs no longer repeats the later paragraphs inside itself: its structured code and result carry `blockCount` (0: not tracked). Read those paragraphs from the story, not `blocks`. Rust code building `StructuredFieldContent` must add the new `block_count` member.
