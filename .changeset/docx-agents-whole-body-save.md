---
"@betteroffice/docx": patch
---

Saving after inserting paragraphs and adding a section break now keeps untouched content exactly as it was, including cell shading, connectors and drawings. If the section boundaries cannot be placed safely, the save is refused instead of rewriting the document body.
