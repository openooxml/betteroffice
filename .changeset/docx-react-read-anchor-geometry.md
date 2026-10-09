---
"@betteroffice/docx-react": minor
"@betteroffice/docx": patch
---

Adds plugin `geometry.readAnchorGeometry`, which resolves paragraph, search, range and revision targets in the default worker viewers and lists pages not yet built in `unbuiltPages`. It refuses with `layout-unavailable` until the pages paint. On by default.
