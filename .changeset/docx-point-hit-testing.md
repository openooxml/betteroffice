---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

Expose `RenderedDomContext.getPositionAtPoint()` for querying a client point without changing selection or focus, including page and story-region identity. The member is optional, so custom `RenderedDomContext` implementations written for earlier versions keep compiling; plugin and editor point queries answer `null` through a context that omits it.

Add `DocxEditorRef.getPositionAtPoint()` and `DocxPluginGeometry.getPositionAtPoint()`, which return the hit with the document `version` its layout shows and a collapsed accepted-view `target` for edit batch steps, and `null` while input is pending or the painted layout is behind the document.

Accept the hit result in `PagedEditorRef.displayPositionToYrsLoc()` to map body, header, footer, and note positions into their editing locations.
