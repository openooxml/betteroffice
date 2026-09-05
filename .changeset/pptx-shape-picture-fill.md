---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Paint `a:blipFill` on a shape: a stretched picture fill now resolves its blip and reaches the display list as an image masked by the shape's own outline, honouring `a:srcRect` and the `a:stretch/a:fillRect` band. Tiled fills are unchanged.
