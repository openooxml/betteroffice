---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Shapes styled with a theme effect reference now draw the theme's outer shadow, as PowerPoint does. Sessions stored by earlier versions recover these shadows when their source is attached, without changing the stored format.
