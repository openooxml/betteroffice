---
'@betteroffice/rust-crates': minor
'@betteroffice/docx-native': minor
---

Adds `prime_last_slice_cost`, what the last `prime_save_seed_states_slice` call charged against its byte budget, story and segment costs included, to the wasm and native DOCX edit sessions. It reads 0 after a slice that charged nothing.
