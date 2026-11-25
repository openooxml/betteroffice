---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': minor
---

Editing beside text-anchored floats respects section grid changes and can reuse unchanged layout when wrapping is independent of page placement. Rust adds `FloatFlowStep`, `FloatFlowReplay`, and `FloatSegmentFlow`, with opt-in recording through `FloatFlow::record_steps` and retrieval through `FloatFlow::steps` or `FloatFlow::take_steps`.
