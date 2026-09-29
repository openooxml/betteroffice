---
"@betteroffice/docx-react": patch
---

`getAnchorGeometry` reads the document once per version and reuses those reads for later targets, so anchoring many proposals no longer rereads the whole document for each one.
