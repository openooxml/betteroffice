---
"@betteroffice/vsdx": patch
"@betteroffice/vsdx-react": patch
"@betteroffice/vsdx-i18n": patch
"@betteroffice/rust-crates": patch
---

Select several shapes at once and keep a compound gesture on one undo entry. Undo, redo, Delete, Escape and the arrow keys now work while any editor chrome holds focus, Ctrl+A selects every visible shape on the page, and nudges, deletes and ribbon colour, rotate and flip commands apply to the whole selection. Core adds `moveShapes`, `deleteShapes` and `setCellFormulas`, each a single transaction that refuses the whole batch with no partial write when one shape's mutation policy refuses. Rotation now goes through a `Rotate` gesture, so `LockRotate` refuses it and the rotation grip disappears on shapes that cannot rotate.
