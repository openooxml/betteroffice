---
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
---

`SaveReport` gains `whole_model_write`, true when a save rewrote every part from the model. Structure the model does not carry may then be gone, and `structure_dropped` is not computed, so `None` there no longer reads as nothing dropped.
