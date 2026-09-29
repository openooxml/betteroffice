---
"@betteroffice/docx-react": patch
---

Add `whenLayoutComplete()` to the editor ref: it resolves with the page count once the whole document is laid out and ready to paint, not just the first pages a large document paints before the rest. `getTotalPages()` now reports 0 until then instead of 1, and a settled display list never shows a layout of part of the document. The DOCX quality harness waits for this signal, so captures of large documents include every page again instead of the first one.
