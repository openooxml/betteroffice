---
'@betteroffice/pptx': patch
---

Presentations with EMF or WMF pictures no longer download the picture converter again on every edit when it fails to load; retries back off and stop after a few attempts.
