---
'@betteroffice/xlsx-react': patch
'@betteroffice/xlsx-i18n': patch
---

Clipboard access is guarded instead of assuming `navigator.clipboard`. On insecure origins or browsers without the API, copy, cut and paste report a localized clipboard refusal and keep accepted cell input; paste no longer surfaces as `input-failed`.
