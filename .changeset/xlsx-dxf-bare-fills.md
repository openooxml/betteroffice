---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

A conditional format whose fill names two different colors paints no fill, and a rule the calc engine cannot evaluate in full paints nothing. Rust: `DifferentialFormat` gains `is_released_reading_of`, `Stylesheet` gains `dxf_fill_colors_differ`, and `xlsx_calc::EvalContext::has_unhandled_unsupported_function` is public.
