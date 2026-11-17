---
'@betteroffice/docx': patch
---

Resolving a caret or selection position in a document with many tables finds the story holding it directly instead of scanning every story, which made each keystroke slower as the document grew.
