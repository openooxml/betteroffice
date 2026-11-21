---
'@betteroffice/docx-react': patch
---

Scrolling a document no longer re-renders the toolbar and editor shell for every painted frame, which lowers main-thread work while scrolling. The page indicator also stays visible while frames arrive during a scroll.
