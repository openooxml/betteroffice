---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Typing in the worker-only docx editor paints each keystroke on the caret's line before the worker replies, drawing the layout engine's own glyph outlines at its pen positions, and then reconciles to the reply. The echo declines and waits for the reply whenever it cannot match the reply's pixels: characters outside printable ASCII and common punctuation, lines that are not left-aligned left-to-right text in one primary font, character pairs the font substitutes, fonts with contextual positioning, edits that change the line's wrap boundaries or height, and pages not yet re-rastered after a zoom or pixel ratio change. `@betteroffice/docx/layout/render` adds `drawGlyphRunPrimitive` and an optional `clip` argument to `rasterizeDisplayPageToBackBuffer`.
