---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Undoing a freeze pane that won a concurrent edit no longer fails: every peer shows the source file's pane and redo brings the winner back. Peers on older versions refuse that undo's update, so update them together.
