---
'@betteroffice/pptx': patch
'@betteroffice/python-pptx': patch
'@betteroffice/rust-crates': patch
---

Draw the shape a Wingdings, Wingdings 2, Wingdings 3 or Webdings `a:buChar` addresses. Those faces reach their glyphs by font position, through a private-use cmap at `U+F0xx` or the raw byte, so the character a deck stores for one of them names a slot rather than the character to draw. Slots for circles, squares, hollow boxes, diamonds, and the solid triangles and arrowheads now resolve to the nearest Unicode character the bundled faces cover. The shapes come from the faces' own glyph names and the sizes from matching each glyph's ink width against the candidates', confirmed against PowerPoint's own render of the corpus decks. An `a:buChar` under a text face, or one that is already a real Unicode character, is drawn exactly as authored.
