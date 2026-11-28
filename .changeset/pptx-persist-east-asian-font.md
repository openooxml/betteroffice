---
"@betteroffice/pptx": patch
"@betteroffice/pptx-native": patch
"@betteroffice/python-pptx": patch
"@betteroffice/rust-crates": patch
---

A font picked for Chinese, Japanese or Korean text in a presentation now writes the DrawingML East Asian typeface (`a:ea`) on save, so PowerPoint keeps the chosen font on reopen. Latin-only edits and non-font changes preserve the authored East Asian typeface.
