---
"@betteroffice/docx": patch
---

Paged structured exports of documents whose footnote placement alternates between layouts now succeed instead of returning `layout-not-converged`, with a `note-layout-fallback` diagnostic on each page that keeps extra note space.
