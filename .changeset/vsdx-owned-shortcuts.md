---
"@betteroffice/vsdx-react": patch
---

Handle Ctrl+S anywhere in the editor: it downloads the diagram instead of letting the browser save
the page, including while a shape's text is being edited. Closing or reloading the tab now warns
while there are unsaved changes or an open text draft.
