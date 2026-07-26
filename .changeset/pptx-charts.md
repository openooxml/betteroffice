---
"@betteroffice/pptx": patch
"@betteroffice/pptx-react": patch
"@betteroffice/rust-crates": patch
---

Charts in a presentation render for real instead of drawing a grey placeholder. Chart parts are loaded through the slide, layout and master relationship cascade, their colours resolve against the deck theme, and the plot streams into slide primitives with an accessible label. Data labels, axis titles and per-point colours draw, and an `ofPie` group now plots as a pie rather than as columns.
