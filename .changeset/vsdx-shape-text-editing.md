---
"@betteroffice/vsdx": patch
"@betteroffice/vsdx-react": patch
"@betteroffice/vsdx-i18n": patch
"@betteroffice/rust-crates": patch
---

Add in-place shape text editing to the VSDX editor. Text edits run through the mutation policy with a typed receipt, are undoable and authorized on remote updates, and saving patches only the edited shape's `Text` element.
