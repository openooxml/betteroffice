---
"@betteroffice/docx": minor
---

**Breaking:** Code that implements `YrsSession` from `@betteroffice/docx/yrs` must add its new members; `insertText` and `deleteRange` now return `YrsReplaceReceipt`, which extends `YrsRevisionReceipt`.
