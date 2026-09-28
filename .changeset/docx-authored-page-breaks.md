---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Keep authored page breaks (`w:br w:type="page"`) when a document is saved while their paragraph's text is unchanged; they were dropped from the saved file, and a selective save of a paragraph holding one failed with "selective document patch is unsafe". A page break inside a paragraph whose text was edited is still dropped on save.
