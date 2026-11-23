---
'@betteroffice/docx': minor
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': minor
---

One run holding two or more whole fields keeps its fields and the text between them on save while the field is unedited: `ComplexField` gains optional `sourceRunXml` and `sourceRunDigest` (`source_run_xml`, `source_run_digest` in the Rust struct; struct literals must add both).
