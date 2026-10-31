---
"@betteroffice/rust-crates": patch
"@betteroffice/pptx": patch
---

Saving an edited slide keeps the original XML of every shape and element the edit did not change, including attribute order, namespace declarations, comments and processing instructions, instead of rewriting the whole slide.
