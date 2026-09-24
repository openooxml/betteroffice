---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

Expose `RenderedDomContext.getPositionAtPoint()` for querying a client point without changing selection or focus, including page and story-region identity.

Accept the hit result in `PagedEditorRef.displayPositionToYrsLoc()` to map body, header, footer, and note positions into their editing locations.
