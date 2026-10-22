---
"@betteroffice/xlsx": patch
---

A long recalculation no longer ends the workbook session after 60 seconds: the engine reports progress while it works, so the 60-second watchdog only fires for a worker that has really gone silent, and edits queued behind the recalculation apply once it returns.
