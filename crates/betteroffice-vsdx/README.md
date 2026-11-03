# betteroffice-vsdx

Typed native facade for inspecting VSDX diagrams.

PNG rendering is enabled by default. Set `default-features = false` to omit
raster support. Targets use zero-based page indexes and the default scale is 1.0.

```rust
use betteroffice_vsdx::{Diagram, PageTarget, RenderOptions};

let diagram = Diagram::open(&vsdx_bytes)?;
let png = diagram.render_png(PageTarget { page: 0 }, &RenderOptions::default())?;
```
