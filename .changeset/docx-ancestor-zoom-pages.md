---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Keep the visible pages painted when a host scales the editor with CSS `zoom` on an ancestor. The page window compared zoomed client rects with unzoomed page offsets, so at any ancestor zoom other than 1 it released the pages on screen and left them blank. The page window, `scrollToPage` and the other scroll-to APIs, scroll anchoring across relayouts, caret scrolling, and the table-resize and remote-presence page windows now convert client pixels to layout pixels through the element's CSS zoom. Scroll-to and anchoring also measure the window, not the document element, when the page itself scrolls. `effectiveZoom(element)` is exported from `@betteroffice/docx/layout/render` for hosts that do the same conversion.
