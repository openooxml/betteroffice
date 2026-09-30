---
"@betteroffice/docx-react": patch
---

Add `whenLayoutComplete()` to the editor ref, which resolves with the page count once the whole document is laid out, not just its first pages. Until then `getTotalPages()` returns 0 instead of 1.
