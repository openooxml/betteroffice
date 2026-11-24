---
'@betteroffice/rust-crates': patch
'@betteroffice/python-pptx': patch
---

`Presentation::save` in Rust and `Presentation.save` in Python succeed after slides are deleted, and prepare fonts for inserted and duplicated slides.
