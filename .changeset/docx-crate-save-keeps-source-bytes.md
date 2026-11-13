---
'@betteroffice/rust-crates': patch
---

Saving a document opened through `betteroffice-docx` keeps the bytes of every part and paragraph it did not change. An unedited save returns the opened file unchanged; an edited save uses the product's save plan, so tracked table grid changes and other unmodelled content survive.
