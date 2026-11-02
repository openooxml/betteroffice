---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
"@betteroffice/xlsx-native": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/rust-crates": patch
---

The editor and bindings now refuse typing, pasting, clearing or moving locked cells on protected sheets, and formatting or chart moves those sheets do not allow. `editCell` and `editCells` throw for locked cells instead of writing them.
