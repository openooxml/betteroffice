---
'@betteroffice/pptx': patch
'@betteroffice/rust-crates': patch
---

pptx: a slide whose background fill is fully transparent now renders on white paper instead of leaving the page unpainted. A master declaring `solidFill` at `alpha="0"` produced a paint rather than no paint, so the existing white fallback never fired and the slide came out as a hole onto whatever sat behind the canvas — which a PNG export then flattened to black. PowerPoint paints those slides white.
