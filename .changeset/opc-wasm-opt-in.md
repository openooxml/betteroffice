---
"@betteroffice/rust-crates": minor
---

The `betteroffice-opc` `wasm` feature is opt-in instead of default, so native dependants no longer pull wasm-bindgen and js-sys; wasm consumers enable it explicitly.
