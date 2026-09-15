---
"@betteroffice/rust-crates": minor
---

The `betteroffice-opc` `wasm` feature is opt-in instead of default, so the crate itself no longer pulls wasm-bindgen and js-sys unless requested; wasm consumers enable it explicitly.
