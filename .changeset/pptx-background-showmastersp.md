---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Paint slide background pictures and honour `p:sld/@showMasterSp` on the layout. A `p:bg` declaring a `a:blipFill`, or a `p:bgRef` resolving to one through the theme's `a:bgFillStyleLst`, now paints as a full-slide image instead of a flat grey, and `p:bgRef` resolves the referenced fill style rather than only its colour override. A slide that turns off master shapes also drops the layout's own decoration, which is what PowerPoint draws, since the master reaches the slide through the layout.
