---
'@betteroffice/docx': patch
---

Collaboration update listeners now run after each engine call returns instead of inside the document transaction, so a listener can read or edit the session, and a listener that throws no longer stops the others from receiving the update.
