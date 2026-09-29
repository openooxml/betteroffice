---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Find a page's canvas without searching the pages host. The canvas renderer registers its page canvases, and page lookups for pointer routing, the caret, selection and highlight overlays, table handles and viewport anchoring read that registry. Each lookup used to run a selector over a subtree that also holds every page's accessibility mirror, hundreds of thousands of nodes on a long document, several thousand times per keystroke.
