---
"@betteroffice/docx": patch
"@betteroffice/agents": patch
---

Preserves authored font slots and language when edits rewrite or split DOCX runs. Inserted paragraphs that copy a neighbour's formatting retain its authored fonts without adding resolved theme fonts or inherited script slots.
