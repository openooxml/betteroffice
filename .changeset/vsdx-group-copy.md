---
"@betteroffice/vsdx": patch
"@betteroffice/vsdx-react": patch
"@betteroffice/vsdx-i18n": patch
---

Cut, copy, paste and duplicate a shape, and carry a group's whole child tree with it. Core adds `addShapeWithText`, `addShapeTree` and `subtreeGlue`; the React package exports `copySelection`, `pasteEntry` and `duplicateEntry` and puts the commands on the Home tab with their Ctrl shortcuts. A copy carries the glue wholly inside its subtree, remapped onto the copy, and names the reason for any unportable content; a paste applies within the document it was copied from, when every shape it references was copied with it.
