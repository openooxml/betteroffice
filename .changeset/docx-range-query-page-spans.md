---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Answer body range and caret queries from the pages they touch. The display-list query store keeps each page's body position span, read the first time a range query needs it, carried across page reuse and widened across position shifts, and a body range query reads only the pages whose span it meets. Every caret query used to walk every primitive of every page.
