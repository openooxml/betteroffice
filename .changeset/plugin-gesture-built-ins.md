---
'@betteroffice/docx-react': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/pptx-react': minor
'@betteroffice/docx-i18n': patch
'@betteroffice/xlsx-i18n': patch
'@betteroffice/pptx-i18n': patch
---

Allows granted mutating built-in commands during trusted events in a plugin's own UI with `commands` and `document: 'write'` grants. Calls by that handler and microtasks it schedules qualify; event notifications (`onEvent`) and calls outside dispatch refuse with `unsupported-policy`.
