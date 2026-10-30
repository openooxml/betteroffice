---
"@betteroffice/pptx": patch
"@betteroffice/pptx-native": patch
"@betteroffice/rust-crates": patch
---

The same font bytes registered under several family names are now held once and no longer count against the 256-face limit. Opening a deck with more than 256 distinct faces now fails with `PptxResourceLimitError` instead of a closed worker connection.
