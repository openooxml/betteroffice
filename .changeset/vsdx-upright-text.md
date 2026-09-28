---
"@betteroffice/vsdx": patch
---

Keep a shape's text upright when the shape is flipped. `FlipX` or `FlipY` mirrored the glyphs with the outline, and a chain that flipped both axes painted them upside down; the text box now reflects about the shape's own box once per flipped axis, so the label stays unmirrored over the geometry the flip moved, inside flipped groups too.
