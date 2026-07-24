---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Opening a document now seeds the collaborative session directly in the Rust engine instead of materializing the full TypeScript document model and projecting it; the TS model is built lazily only where the public API still exposes it, and the internal DrawingML host package is dissolved.
