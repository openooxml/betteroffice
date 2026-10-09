---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Display lists now clamp to the sheet grid by default. Adds `DisplayTooLargeError` and `getDisplayListCellLimit` so callers can size tiles when a viewport exceeds the display limit.
