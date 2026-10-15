---
"@betteroffice/docx": patch
---

Saving a document with comment replies no longer rewrites the whole body. Replies that already have their own range keep it without duplicates, and a missing reply range is added inside the source paragraph, so text boxes, fields and other content stay intact.
