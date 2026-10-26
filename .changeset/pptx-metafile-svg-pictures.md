---
"@betteroffice/pptx": patch
"@betteroffice/pptx-react": patch
"@betteroffice/rust-crates": patch
---

Metafile pictures with labels, bitmaps or complex paint now use a lazy SVG fallback by default in editors and native PNG exports. Adds `metafileSvg` for resolving cached display bytes while preserving original presentation media when saving.
