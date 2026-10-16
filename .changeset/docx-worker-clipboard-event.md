---
"@betteroffice/docx-react": patch
---

Copy and cut in the worker editor write the selection's text and HTML into the native clipboard event, so they keep working where the async Clipboard API is denied or unavailable.
