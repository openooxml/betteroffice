---
'@betteroffice/docx': patch
---

Saving a document whose comment ids are larger than JavaScript can count exactly no longer hangs when a new comment is added; the new comment takes a free id.
