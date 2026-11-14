---
'@betteroffice/pptx': patch
---

A caret anchor whose slide or shape was deleted now resolves to null, or to its offset again once the deletion is undone. Typing, deleting or formatting in such a story is refused with an error instead of being silently dropped from the saved file.
