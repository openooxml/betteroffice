---
"@betteroffice/docx-react": patch
---

With `experimentalWorkerOpen`, only pages near the viewport are prepared, so memory stays bounded on long documents; `whenLayoutComplete` also waits for the visible pages.
