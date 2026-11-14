---
'@betteroffice/docx-react': patch
---

Scrolling a document no longer repositions the hidden text input on every scroll step; the input now moves with the page content, which lowers main-thread layout work while scrolling.
