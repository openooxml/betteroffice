---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Table resize handles are now built only for pages around the viewport, instead of every table on every page, via the new `deriveDisplayListTableFragmentsOnPages` query.
