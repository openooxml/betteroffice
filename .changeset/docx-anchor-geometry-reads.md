---
"@betteroffice/docx-react": patch
---

`getAnchorGeometry` reads the document's revisions, paragraph spans, story segments and paragraph anchors once per document version and reuses them for every later target at that version. Anchoring a proposal on a 267-page document drops from 170-300 ms per call to a few milliseconds after the first.
