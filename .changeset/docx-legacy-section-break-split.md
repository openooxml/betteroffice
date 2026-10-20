---
"@betteroffice/docx": patch
---

Section breaks inserted through the `legacySyncApi` session now split the paragraph as the worker editor does, so worker replicas apply them and a break in the final section keeps its headers and footers. Page breaks land after a paragraph's leading tables.
