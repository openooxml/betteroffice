---
"@betteroffice/vsdx": patch
"@betteroffice/vsdx-react": patch
"@betteroffice/vsdx-i18n": patch
"@betteroffice/rust-crates": patch
---

Render only shapes on visible layers and add a layers panel with per-layer visibility toggles. Core adds layer-aware rendering with `pageLayers`, `setLayerVisible` and `clearLayerVisibility`. Hiding a layer clears the selection of shapes it hides so no hidden shape keeps handles or commands.
