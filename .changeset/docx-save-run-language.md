---
"@betteroffice/docx": patch
---

Saving a document no longer strips the language set on text the edit did not touch, so Word keeps checking its spelling and grammar in that language. Complex script theme fonts are saved under the attribute name Word reads, so Word applies them.
