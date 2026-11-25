---
'@betteroffice/pptx': patch
---

Presentations with EMF or WMF pictures no longer download the picture converter again on every edit when it fails to load; retries back off. The first frame no longer waits more than 3 seconds for the converter and shows the pictures once they are ready.
