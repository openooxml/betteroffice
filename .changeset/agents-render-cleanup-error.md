---
'@betteroffice/agents': patch
'@betteroffice/docx': patch
---

Agent document operations now report the original failure when cleaning up after it also fails, instead of the cleanup error.
