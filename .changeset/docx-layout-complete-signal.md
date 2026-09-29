---
"@betteroffice/docx-react": patch
---

Add `whenLayoutComplete()` to the editor ref. It resolves with the page count once the whole document is laid out and ready to paint, not just the first pages that a large document paints ahead of the rest. Until then `getTotalPages()` returns 0 instead of 1. A settled display list never shows a partial layout, and the DOCX quality harness now waits for the full layout, so captures of large documents include every page instead of only the first.
