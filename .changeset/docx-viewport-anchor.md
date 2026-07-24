---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Remote collaborators' edits no longer move the local viewport: relayouts triggered by remote updates anchor to the topmost visible line via yrs sticky positions and compensate the scroll offset, while caret scrolling fires only for local actions.
