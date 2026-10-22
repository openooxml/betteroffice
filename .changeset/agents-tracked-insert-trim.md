---
'@betteroffice/agents': patch
---

DOCX proposal replacements now preserve shared text by default, tracking only the changed middle. Adds optional `after` to PPTX `duplicate_slide` to place a copy after a slide ID, mutually exclusive with `index`.
