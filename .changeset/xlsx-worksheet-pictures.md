---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/rust-crates': minor
---

Worksheet pictures now appear in display lists as `image` commands, on by default, and in native PNG exports. `paintDisplayList` draws them when given `resolveImage`, fed by the new `mediaBytes`; otherwise it skips them.
