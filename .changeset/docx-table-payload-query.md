---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Read the current table's structure without reading its whole story. The toolbar's table state, the table target and the table-properties dialog each found the table by listing every segment of the story that holds it, several times per keystroke, which cost hundreds of milliseconds per key in a table cell of a long document. A new `tablePayload` session query returns just that table's payload, the target reads it once, and the dialog reads it only while open.
