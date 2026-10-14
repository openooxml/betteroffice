---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Saving no longer writes inherited formatting into the document. Run, cell and table properties are saved relative to the styles and defaults they inherit, and inserting or deleting table rows keeps the other rows and paragraphs as they were. Toggle properties such as bold now combine across styles as the OOXML specification defines, so documents that set one in two style types, or only in a table style, can render differently than before.
