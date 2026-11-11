---
'@betteroffice/xlsx-react': patch
---

The legacy read-only viewer now reports unsupported plugins again through `onPluginError` when a plugin is removed and re-added, when the file changes, or when the viewer returns after the editor.
