---
"@betteroffice/docx-react": patch
---

Build canvas table resize handles only for the pages around the viewport. The overlay derived table fragments and projected a handle for every table on every page on each display-list update; it now follows the visible page window (the viewport and one page either side, as remote presence does) and updates it as the page scrolls.
