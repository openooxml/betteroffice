---
'@betteroffice/rust-crates': patch
---

Opening a DOCX with `betteroffice-docx-edit` uses less memory by avoiding a duplicate copy of the parsed document.
