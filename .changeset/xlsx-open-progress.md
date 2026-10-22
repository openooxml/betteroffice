---
"@betteroffice/xlsx": patch
---

Opening a large workbook now reports progress while it is read, built and recalculated, so a long open keeps the worker session alive instead of hitting the 60-second watchdog.
