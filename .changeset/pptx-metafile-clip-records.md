---
'@betteroffice/pptx': patch
'@betteroffice/rust-crates': patch
---

pptx: a picture stored as an EMF metafile now draws when the file sets a clip. The metafile player treated any unlisted record as fatal and abandoned the whole drawing, so a single `EXTSELECTCLIPRGN` — record 12 of a typical file, and a no-op where every corpus occurrence resets to the default region — silently blanked the picture. Rectangular clips are now tracked on the device context, survive `SAVEDC`/`RESTOREDC`, and narrow the shape they draw into; a clip the player cannot represent still refuses the drawing, so nothing paints unclipped. 18 of 69 corpus metafiles decode where 10 did before, and on the 103-deck fidelity corpus `pptarena-042` moves 0.7229 to 0.8451 with no deck regressing.
