---
'@betteroffice/docx-react': patch
---

In the worker-only docx editor, pages that scroll into or out of view build or trim their accessibility mirror at idle time once scrolling settles, instead of during the scroll.
