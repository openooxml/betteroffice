---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Build canvas table resize handles only for the pages around the viewport. The overlay derived table fragments and projected a handle for every table on every page on each display-list update; it now follows the visible pages and one page either side, updated as the page scrolls or its scroll container resizes. `deriveDisplayListTableFragmentsOnPages` derives the fragments of a page range with each table's identity still taken from every page it covers.
