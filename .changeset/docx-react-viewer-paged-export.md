---
"@betteroffice/docx-react": patch
"@betteroffice/docx": patch
---

In viewer sessions `exportStructuredWithPages` reads from the document worker without a main-thread document copy, and its result now includes comment authors and dates and source page-break positions.
