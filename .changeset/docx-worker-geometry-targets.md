---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

With `experimentalWorkerOpen`, plugin `getAnchorGeometry` returns the same rects as the default mode for paragraph, search, range and revision targets while the worker holds host proposals in a read-only editor. A target's first request can answer `layout-unavailable` until the layout change that follows.
