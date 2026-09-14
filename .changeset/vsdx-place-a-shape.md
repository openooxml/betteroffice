---
"@betteroffice/vsdx-react": patch
---

Drag a shape from the gallery onto the canvas and it lands where it was dropped, already selected.
A tile click still inserts at the page centre and cascades there, per page, so repeats no longer
stack. Every master inserts at its own aspect ratio, one inch tall, and a right-click that misses
every shape now drops the selection along with the menu it closes.
