---
"@betteroffice/agents": patch
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

DOCX saves keep untouched paragraphs byte for byte after agent and editor edits, across spanning comment ranges, bookmarks, content controls, merged cells and wide tables. Footnote edits keep a direct automatic color, and agent export refuses saves that would rewrite unedited paragraphs.
