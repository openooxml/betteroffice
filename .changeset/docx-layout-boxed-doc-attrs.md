---
"@betteroffice/rust-crates": minor
---

**BREAKING (Rust consumers):** public `DocAttrs` fields in `betteroffice-docx-layout` are now `Option<Box<T>>`; construct them with `Some(Box::new(value))`. The JS API and the wire format are unchanged.
