---
'@betteroffice/pptx-react': patch
'@betteroffice/docx-react': patch
'@betteroffice/xlsx-react': patch
---

Toolbar overflow and context menus now load their menu list separately and preload it when the browser is idle, so the editors' initial download is smaller. A menu that cannot load stays closed and retries.
