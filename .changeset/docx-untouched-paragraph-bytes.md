---
"@betteroffice/agents": patch
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

DOCX saves keep untouched paragraphs byte for byte after agent and editor edits, across spanning comment ranges, bookmarks, content controls, merged cells and wide tables. Footnote text edits keep a direct automatic color and the original run boundaries instead of re-serializing the paragraph.
