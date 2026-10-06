---
"@betteroffice/docx-react": patch
"@betteroffice/xlsx-react": patch
"@betteroffice/pptx-react": patch
---

Fixes toolbar overflow crashes when groups measure differently while hidden, such as when a Content Security Policy blocks the package stylesheet. The overflow menu now settles instead of failing with "Maximum update depth exceeded".
