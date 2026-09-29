---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

A page whose positions all shift by one delta now ships as a single run instead of one per primitive, cutting sync payloads from megabytes to kilobytes.
