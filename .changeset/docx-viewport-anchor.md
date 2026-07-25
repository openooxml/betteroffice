---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Remote collaborators' edits no longer move the local viewport: relayouts triggered by remote updates anchor to the topmost visible line via yrs sticky positions and compensate the scroll offset, while caret scrolling fires only for local actions. Anchoring holds across page boundaries too, so text overflowing onto a new page (or pulling back off one) no longer jumps the viewport for either the author or a viewer.
