---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Keep the visible pages painted when a host scales the editor with CSS `zoom` on an ancestor. The page window compared zoomed client rects with unzoomed page offsets, so at any ancestor zoom other than 1 it released the pages on screen and left them blank. The window, `scrollToPage` and the other scroll-to APIs, scroll anchoring across relayouts and caret scrolling now convert client pixels to layout pixels. `renderedScale(element)` is exported from `@betteroffice/docx/layout/render` for hosts that do the same conversion.
