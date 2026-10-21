---
"@betteroffice/docx": patch
---

Moving the caret or changing the selection in a long document builds the next frame from the pages and part geometry already sent, instead of re-encoding the whole document.
