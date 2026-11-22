---
"@betteroffice/fonts": patch
"@betteroffice/docx": patch
"@betteroffice/pptx": patch
"@betteroffice/vsdx": patch
"@betteroffice/rust-crates": minor
---

Large bundled font faces decode outside the PPTX worker's engine, and registering a font in WebAssembly copies its bytes once, which lowers worker memory for CJK presentations. Rust callers must replace `EditSession::register_measure_font` with `register_owned_measure_font` and pass a `Vec<u8>`.
