---
"@betteroffice/docx-react": patch
---

Adds `whenLayoutComplete()` to the editor ref, resolving with the page count once the whole document is laid out. Await it, not `onFirstPagePainted`, before reading page counts or content: until then `getTotalPages()` returns 0 and `getPageContent(n)` returns `null`, even for painted pages.
