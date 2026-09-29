---
"@betteroffice/rust-crates": patch
"@betteroffice/docx": patch
---

Shrink every retained display primitive by boxing the attributes most primitives never set (content-control, comment, revision, field, note, table and cell metadata, clip groups, leader glyphs, highlight slices, content frames and DrawingML paint values), and trim each page's primitive list to its length. A primitive drops from 1,888 to 728 bytes in wasm, so a fully built display list of a 267-page document holds hundreds of megabytes less in the resident worker. Frames and display-list JSON are byte-identical.
