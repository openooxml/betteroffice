---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Keep authored page breaks (`w:br w:type="page"`) when a document is saved and the runs around them are unchanged; they were dropped from the saved file, and a selective save of a paragraph holding one failed with "selective document patch is unsafe".
