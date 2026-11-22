---
"@betteroffice/docx-native": patch
"@betteroffice/pptx-native": patch
"@betteroffice/docx": patch
---

DOCX and PPTX native sessions can release their observers and memory after an engine failure. The DOCX worker reports failures promptly and passes along queued collaboration updates.
