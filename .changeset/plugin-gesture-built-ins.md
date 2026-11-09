---
'@betteroffice/docx-react': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/pptx-react': minor
'@betteroffice/docx-i18n': patch
'@betteroffice/xlsx-i18n': patch
'@betteroffice/pptx-i18n': patch
---

Allows granted mutating built-in commands during trusted events in a plugin's own UI. Requires `commands` and `document: 'write'` grants. Calls outside the event refuse with `unsupported-policy`.
