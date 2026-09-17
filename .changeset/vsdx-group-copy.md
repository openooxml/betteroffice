---
"@betteroffice/vsdx": patch
"@betteroffice/vsdx-react": patch
"@betteroffice/vsdx-i18n": patch
"@betteroffice/rust-crates": patch
---

Cut, copy, paste and duplicate a shape, and carry a group's whole child tree with it. Core adds `addShapeWithText`, `addShapeTree` and `subtreeGlue`; the React package exports `copySelection`, `pasteEntry` and `duplicateEntry` and puts the commands on the Home tab with their Ctrl shortcuts. Glue wholly inside a copied subtree is remapped onto the copy, glue crossing the copy boundary is dropped, a copy refuses unportable content with its reason, and a paste is refused when it came from another document or would leave a reference to a shape it did not copy.
