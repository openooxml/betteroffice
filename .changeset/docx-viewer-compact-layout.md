---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Read-only sessions take each layout update from the worker as a compact summary instead of the full layout, which cuts main-thread work per update with identical rendering.
