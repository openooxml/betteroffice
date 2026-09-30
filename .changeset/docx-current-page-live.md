---
"@betteroffice/docx-react": patch
---

`getCurrentPage()` now returns the page on screen at the moment it is called, and a read-only editor no longer scrolls back to its caret while pages build, so `scrollToPage` lands on its page.
