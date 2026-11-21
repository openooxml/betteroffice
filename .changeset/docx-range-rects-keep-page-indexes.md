---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Keystrokes in long documents spend less time on selection and paragraph geometry: pages whose text did not change, or only moved after the edit, keep their range index and caret stops between frames instead of indexing every built page again.
