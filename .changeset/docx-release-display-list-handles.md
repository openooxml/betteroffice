---
'@betteroffice/docx': patch
---

The editor closes its display-list store handles when it unmounts and releases worker selection reads when a session ends, instead of leaving them to garbage collection.
