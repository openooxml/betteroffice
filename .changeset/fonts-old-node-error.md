---
'@betteroffice/fonts': patch
---

Reading bundled fonts from installed packages on Node.js older than 20.16 now fails with a clear error and a one-time warning instead of a silent failed `fetch` of a `file:` URL.
