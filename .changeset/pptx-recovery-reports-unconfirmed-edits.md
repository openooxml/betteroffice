---
'@betteroffice/pptx-react': patch
'@betteroffice/pptx': patch
---

When a presentation is recovered after a worker failure, edits made just before the failure whose outcome is unknown are reported through onError and in the recovery result instead of being left out silently.
