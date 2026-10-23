---
'@betteroffice/docx': patch
'@betteroffice/docx-react': patch
---

Fixes Backspace after tables leaving an invalid caret and Backspace between tables deleting table content. Backspace now moves into the previous table's last cell, and forward Delete preserves table boundaries.
