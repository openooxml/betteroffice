---
"@betteroffice/xlsx": patch
"@betteroffice/pptx": patch
---

A long worker call, such as recalculating a large workbook, no longer ends the session after 60 seconds: a call the worker has started may run up to 10 minutes, and edits queued behind it apply once it returns.
