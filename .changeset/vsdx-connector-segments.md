---
"@betteroffice/vsdx": patch
"@betteroffice/vsdx-react": patch
"@betteroffice/rust-crates": patch
---

Edit a connector's filed route. Core adds `setConnectorRoute`, which rewrites a 1-D connector's `Geometry` rows from scene-space points through the mutation policy, and the page display list now reports per-connector selection chrome with the glue kind of each endpoint. The React editor paints one handle per straight segment and drags a segment perpendicular only, re-plumbing the rest of the route.
