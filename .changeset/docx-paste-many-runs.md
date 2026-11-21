---
'@betteroffice/docx': patch
'@betteroffice/docx-react': patch
---

Pasting text with more than 300 formatted runs no longer stalls the editor: such a paste inserts its text without run formatting, as earlier releases did, and offers no Paste Options.
