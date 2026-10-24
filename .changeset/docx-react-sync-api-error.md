---
"@betteroffice/docx-react": patch
---

Exports `DocxSyncApiError` (`code: "sync-api-unavailable"`), which synchronous ref members throw on the default editor, so hosts can check for it with `instanceof`.
