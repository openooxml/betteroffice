---
'@betteroffice/rust-crates': patch
'@betteroffice/xlsx': patch
---

Pictures and shapes stored inside an alternate-content wrapper now follow XLSX row and column edits too, every branch alike. A wrapper whose branches sit on different cells or hold anything but one anchor still refuses the edit.
