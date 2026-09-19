---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Wrap a `start` tab together with the word it cannot fit. A tab landing on a stop that leaves the following word no room used to stay on the line while that word wrapped to the paragraph indent, which cost a line against Word on every such paragraph. The tab now moves to the next line with its word, as Word lays it out. `end`, `center` and `bar` stops already reserve the following content and are unaffected.
