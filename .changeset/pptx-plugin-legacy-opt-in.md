---
"@betteroffice/pptx-react": minor
---

`plugins`, `pluginGrants` and `onPluginError` require the deprecated `legacySyncApi` opt-in. The default worker editor reports unsupported plugins through `onError` instead of silently ignoring them.
