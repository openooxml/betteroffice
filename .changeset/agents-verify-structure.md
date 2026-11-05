---
'@betteroffice/agents': patch
---

`office_verify` on DOCX now compares the saved file with the original: parts, element counts, field instructions, drawings and relationships. Losses the edits do not explain are listed and set `structure.ok` to false. Pass `structure: false` to skip it.
