---
'@betteroffice/docx': patch
'@betteroffice/agents': patch
---

Saving a document no longer hangs when a new comment's id would have to follow a source comment id too large to extend; the save fails with an error saying so. Every other document saves its comment ids as before.
