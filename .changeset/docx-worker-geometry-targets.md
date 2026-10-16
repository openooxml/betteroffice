---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Plugin `getAnchorGeometry` now supports paragraph, search, range and revision targets in the default worker-only editor and in read-only `legacySyncApi` worker viewers that serve host proposals; the first request can return `layout-unavailable` until geometry arrives.
