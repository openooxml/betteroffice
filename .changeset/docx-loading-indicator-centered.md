---
"@betteroffice/docx-react": patch
---

Keep the loading indicator centered over the whole editor until the first page paints. It no longer jumps to the top of the page area when the toolbar and ruler mount, and a custom `loadingIndicator` now shows in that state too.
