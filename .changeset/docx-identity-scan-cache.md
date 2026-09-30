---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Resolving paragraph anchors and reading paragraph identities reuse one read of the document until it changes, so proposing many edits to a long document no longer rereads it for every anchor.
