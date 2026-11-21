---
'@betteroffice/xlsx-react': patch
---

The legacy read-only worker viewer (`experimentalWorkerOpen` with `legacySyncApi` and `readOnly`) reports unsupported plugins again through `onPluginError`, or `console.error` without it, after re-adding, a revision change, a file change, or a return from the editor.
