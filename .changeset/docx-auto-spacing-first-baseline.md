---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

Hang the baseline of an `auto`-ruled line off the top of its box, as Word does. The painter centred the leading a line-spacing multiple adds, so every line of a 1.15- or double-spaced paragraph sat half that leading below Word's baseline; the box height and the pagination it drives are unchanged.
