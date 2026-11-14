---
'@betteroffice/docx-native': patch
---

PNG export on the native engine reads the document's original bytes directly instead of first materializing the caller's session.
