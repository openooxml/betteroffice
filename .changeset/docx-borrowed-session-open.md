---
'@betteroffice/rust-crates': minor
---

Rust callers of `EditSession::seed_from_docx`, `EditSession::open_docx`, `EditSession::open_docx_preview` and `EditSession::open_docx_preview_with_budget` must use the same methods on `EditSessionCore` and handle `EngineError`, or pass owned bytes to the corresponding `EditSession` methods ending in `_owned`.
