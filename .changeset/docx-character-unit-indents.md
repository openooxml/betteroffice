---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Character-unit paragraph indents (`w:leftChars`, `w:rightChars`, `w:firstLineChars`, `w:hangingChars`) now survive a save in the direction they were authored, and an explicit `w:firstLine="0"` is no longer discarded. Documents that indent by characters, as Chinese typography does, previously kept only the twip approximation, a first-line indent could come back as a hanging one, and paragraphs that cancelled an inherited first-line indent silently regained it.
