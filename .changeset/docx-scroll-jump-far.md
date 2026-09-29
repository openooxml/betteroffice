---
"@betteroffice/docx-react": patch
---

Jump instead of animating when `scrollToPage` or `scrollToPosition` targets something more than two viewports away. Smooth-scrolling across a long document painted every page on the way; `scrollToPage(200)` in a 267-page document took about 1.3 s to arrive.
