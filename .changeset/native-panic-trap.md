---
"@betteroffice/docx-native": patch
"@betteroffice/pptx-native": patch
"@betteroffice/xlsx-native": patch
"@betteroffice/docx": patch
"@betteroffice/pptx": patch
"@betteroffice/xlsx": patch
---

A Rust panic in a native engine now fails its session as a trap, as on WebAssembly, instead of aborting the process. The engine object that panicked refuses every later call except free.
