---
"@betteroffice/docx": minor
"@betteroffice/docx-native": minor
"@betteroffice/docx-react": patch
"@betteroffice/docx-i18n": patch
"@betteroffice/rust-crates": minor
---

Screen reader page mirrors read a visible list bullet as a localised word, never its stored character. `DocAttrs` gains `listMarkerBullet` and `MirrorLabels` gains `bullet`; without either, a bullet stays hidden. Rust `DocAttrs` literals must set `list_marker_bullet`.
