---
"@betteroffice/pptx": patch
---

The worker editor opens in Next.js (Turbopack) apps again. It fetched its engine from the asset's href as given, which Turbopack makes a bare path, so every editor failed with `Invalid URL`.
