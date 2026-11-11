---
'@betteroffice/docx-react': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/pptx-react': minor
'@betteroffice/docx-i18n': patch
'@betteroffice/xlsx-i18n': patch
'@betteroffice/pptx-i18n': patch
---

Granted mutating built-ins now run from a trusted gesture in the plugin's own UI until that event's dispatch and queued microtasks finish; other plugins and later calls refuse with `unsupported-policy`. Plugin panels, overlays and sidebar cards sit in a `display: contents` wrapper.
