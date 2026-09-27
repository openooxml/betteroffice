---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

Expose `RenderedDomContext.getPositionAtPoint()` for querying a client point without changing selection or focus, including page and story-region identity.

Add `DocxEditorRef.getPositionAtPoint()` and `DocxPluginGeometry.getPositionAtPoint()`, which return the hit with the document `version` its layout shows and a collapsed accepted-view `target` for edit batch steps, and `null` while input is pending or the painted layout is behind the document.

Accept the hit result in `PagedEditorRef.displayPositionToYrsLoc()` to map body, header, footer, and note positions into their editing locations.
