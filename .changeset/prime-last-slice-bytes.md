---
'@betteroffice/rust-crates': minor
'@betteroffice/docx-native': minor
---

Adds `prime_last_slice_bytes`, the bytes the last `prime_save_seed_states_slice` call digested, to the wasm and native DOCX edit sessions. It reads 0 after a slice that digested none.
