---
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
---

`SaveReport` gains `whole_model_write`, true when a save had no session plan and wrote the stories from the model. Structure the model does not carry may then be gone, and `structure_dropped` is not computed, so its `None` no longer reads as nothing dropped.
