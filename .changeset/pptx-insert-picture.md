---
"@betteroffice/pptx": minor
"@betteroffice/pptx-react": minor
"@betteroffice/pptx-i18n": minor
"@betteroffice/rust-crates": minor
---

Insert a picture onto a slide from the editor. The image mints its own media part, content-type default and relationship on save; `PptxEditor` gains a small "Insert image" icon button next to the text-box tool, and `PresentationHandle` gains `addPicture`. Unsupported MIME types and images over 8 MiB are rejected before the picture reaches the deck, keeping oversized bytes out of collaboration updates.
