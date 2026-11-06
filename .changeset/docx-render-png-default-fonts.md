---
'@betteroffice/docx': patch
---

`renderPng` with no `fontProvider` gives fonts the bundled set doesn't carry the last-resort face in the browser, as the native engine does, instead of estimated metrics.
