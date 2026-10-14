---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Typing in the worker-only docx editor paints each keystroke before the worker replies, in the layout engine's own glyphs, and waits for the reply when it cannot match it exactly. Adds `drawGlyphRunPrimitive` and an optional `clip` argument to `rasterizeDisplayPageToBackBuffer`.
