---
"@betteroffice/docx": patch
---

Saving a document now keeps table rows that skip leading or trailing grid columns (`w:gridBefore`, `w:gridAfter`) together with their widths, so those rows keep their position and shape in Word.
