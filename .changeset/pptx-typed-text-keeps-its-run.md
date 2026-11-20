---
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
'@betteroffice/python-pptx': patch
'@betteroffice/rust-crates': patch
---

Text typed with a run's own formatting stays in that run when the run says `strike="noStrike"`, instead of starting a new run without it. Typing after turning strikethrough off still comes out unstruck.
