---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Character-unit paragraph indents (`w:leftChars`, `w:rightChars`, `w:firstLineChars`, `w:hangingChars`) now survive a save, and an explicit `w:firstLine="0"` is no longer discarded. Documents that indent by characters, as Chinese typography does, previously kept only the twip approximation, and paragraphs that cancelled an inherited first-line indent silently regained it.
