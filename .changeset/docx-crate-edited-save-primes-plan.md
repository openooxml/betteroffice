---
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
---

An edited save from the Rust crate or the Python package no longer seeds the source document twice when it has at most 4,096 paragraphs and 16 MiB. Median edited save on the public documents: 38.1 ms to 25.6 ms. Output bytes are unchanged.
