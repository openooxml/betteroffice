---
"@betteroffice/vsdx": patch
"@betteroffice/rust-crates": patch
---

Migrate VSDX diagrams stored under the previous CRDT schema. Stories written before in-place text editing held resolved text tokens as JSON under the same schema version as the plain text that replaced them, so reopening such a document showed and saved a JSON blob as every shape's text. The schema version now distinguishes the two and a stored document is carried forward when it is opened.
