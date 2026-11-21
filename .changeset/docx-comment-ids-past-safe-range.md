---
'@betteroffice/docx': patch
'@betteroffice/agents': patch
---

Saving a document whose comment ids are larger than JavaScript can count exactly no longer hangs when a new comment is added: the new comment takes the next free id, and a save with no free id left fails with an error.
