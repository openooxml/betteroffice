---
'@betteroffice/docx': minor
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': minor
---

The DOCX engine holds list definitions the editor creates and writes them on save. Documents without one are unchanged. New in the edit crate: `ListKind`, `ListIdentity`, `CreatedListDefinition`, `resolve_list_identity`, `ensure_list_definition` and `created_list_definitions`.
