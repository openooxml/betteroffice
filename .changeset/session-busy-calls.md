---
"@betteroffice/xlsx": patch
"@betteroffice/pptx": patch
---

A worker call that runs for a long time, such as recalculating a large workbook, keeps the session open: the worker announces each call as it starts, and the client allows a started call up to 10 minutes instead of 60 seconds of silence. Edits queued behind it apply once it returns.
