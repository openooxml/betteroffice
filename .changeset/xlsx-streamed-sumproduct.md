---
"@betteroffice/xlsx": patch
---

SUMPRODUCT over long ranges, including whole columns and ranges declared past the used rows, recalculates with far fewer cell reads, so large summary sheets finish within the recalculation budget instead of showing `#NUM!`.
