---
"@betteroffice/xlsx": minor
"@betteroffice/xlsx-react": patch
"@betteroffice/xlsx-native": minor
"@betteroffice/python-xlsx": minor
"@betteroffice/rust-crates": minor
---

Editors and direct writes now refuse locked cells, formats and chart moves and resizes on protected sheets by default, like Excel. Direct writes accept opt-in `protectedSheets: "allow"` (Python `protected_sheets="allow"`), which writes through and keeps the protection.
