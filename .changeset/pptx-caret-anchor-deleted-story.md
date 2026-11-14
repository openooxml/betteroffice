---
'@betteroffice/pptx': patch
---

A caret anchor whose slide or shape was deleted now resolves to null, as documented, instead of its old offset. It resolves to its offset again if the deletion is undone.
