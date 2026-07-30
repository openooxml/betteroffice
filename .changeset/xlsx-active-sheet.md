---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

The active tab now survives a round trip. Opening a workbook reads `activeTab` from the first `workbookView` rather than always starting on the first sheet, and saving writes the selection back — patched into the preserved `bookViews` where the source has one, emitted where it does not. An `activeTab` past the last sheet falls back to the first, and a save that leaves the active sheet alone still returns the source bytes untouched.
