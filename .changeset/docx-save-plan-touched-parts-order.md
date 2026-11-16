---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
"@betteroffice/rust-crates": patch
---

The DOCX paragraph save plan now lists `touchedParts` sorted by part path, with each part's paragraph ordinals ascending, so the same edits give the same plan in every session. Saved files are unchanged.
