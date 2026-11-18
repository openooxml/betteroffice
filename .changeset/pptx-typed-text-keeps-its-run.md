---
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
'@betteroffice/python-pptx': patch
'@betteroffice/rust-crates': patch
---

Typed text stays in the run it is typed into when that run carries explicit default formatting such as `strike="noStrike"`, instead of starting a new run without it. Typing after turning strikethrough off still comes out unstruck.
