---
'@betteroffice/pptx-react': patch
'@betteroffice/pptx': patch
'@betteroffice/python-pptx': patch
'@betteroffice/rust-crates': minor
---

Drag and resize placeholders that take their geometry from the layout or master, from the frame they are drawn in, and draw them with the rotation and flips they inherit. The snapshot carries that geometry as `inherited`; the first `moveShape`, `resizeShape` or `setShapeRect` makes the whole transform the placeholder's own, so it keeps its size and orientation live and after a save. The Rust crates add `Placeholder::matches`, the placeholder matching that rendering and editing share, and the facade re-exports `InheritedGeometry`.
