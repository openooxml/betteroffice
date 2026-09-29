---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Read a document's revisions and its sidebar projection once per document version. `listRevisions` reads each story once per call, instead of once for every change's range and preview; on a 267-page document with ten suggestions it drops from about 250 ms to about 20 ms. The sidebar projection is reused until the document changes, and the paged editor's sidebar anchors, which it recomputes whenever the display list changes, reuse the revisions and tracked-change entries too. Scrolling or building pages of a document with tracked changes no longer relists them.
