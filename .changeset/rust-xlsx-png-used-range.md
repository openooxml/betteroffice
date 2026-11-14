---
'@betteroffice/rust-crates': minor
---

Adds the opt-in `PngTarget::UsedRange { sheet }` target for XLSX PNG rendering. It renders exactly as `render_sheet` does without a range: the sheet's used range at call time.
