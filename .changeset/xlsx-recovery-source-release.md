---
"@betteroffice/xlsx": patch
---

Spreadsheet sessions release their copy of the original file once crash recovery holds a checkpoint, saving memory equal to the file size. An unreadable checkpoint now fails recovery with a clear `checkpoint-unreadable` error.
