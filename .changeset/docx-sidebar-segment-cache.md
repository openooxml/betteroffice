---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

The sidebar no longer re-reads a long document's unchanged paragraphs after each proposal or edit. `createYrsSidebarProjection` takes an optional `YrsStorySegmentSource` to read story segments through.
