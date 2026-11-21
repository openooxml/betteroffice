---
'@betteroffice/rust-crates': minor
'@betteroffice/docx-native': minor
---

Adds `prime_last_slice_cost` to the wasm and native DOCX edit sessions: what the last `prime_save_seed_states_slice` call charged against its byte budget, overshoot included, or 0 (committed-scan slices charge nothing). The docx worker sizes the next slice from it, growing at most twofold.
