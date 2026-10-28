---
'@betteroffice/pptx': patch
---

Disposing a PowerPoint worker editor while an edit, read or thumbnail call is still pending no longer reports an unhandled promise rejection; the call still rejects for any caller that awaits it.
