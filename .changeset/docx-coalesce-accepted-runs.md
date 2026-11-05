---
'@betteroffice/docx': patch
'@betteroffice/agents': patch
'@betteroffice/rust-crates': patch
---

Accepted DOCX text replacements now coalesce adjacent changed runs with identical written properties by default. Untouched run splits, markers, fields and tracked revisions are preserved.
