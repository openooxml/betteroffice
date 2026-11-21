---
'@betteroffice/fonts': patch
---

Reading bundled fonts from installed packages on Node.js versions without `process.getBuiltinModule` (before 20.16, or 21 to 22.2) now fails with a clear error and a one-time warning instead of a silent failed `fetch` of a `file:` URL.
