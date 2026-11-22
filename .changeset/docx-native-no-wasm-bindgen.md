---
"@betteroffice/rust-crates": minor
---

DOCX crates no longer compile wasm-bindgen on native targets. On native targets, replace `layout_document_json` with `layout_document_json_str`, import `hit_test_json`, `vertical_move_json`, `range_rects_json`, `range_rects_region_json` and `hit_test_regions_json` from `docx_layout::hit`, and enable the opt-in `wasm` feature for `parse_docx_s9` or call `parse_docx_s9_json`.
