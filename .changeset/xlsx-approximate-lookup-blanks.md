---
"@betteroffice/xlsx": patch
---

Fixes approximate lookups returning blank cells as zero or selecting keys of a different type. Approximate matching skips blanks and keys of other types by default while preserving exact-match behavior.
