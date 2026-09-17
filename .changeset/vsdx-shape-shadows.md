---
"@betteroffice/vsdx": minor
---

Render shape shadows. A shape's `ShdwPattern`, offsets and colour resolve through its own cells, its cached values and the theme effect scheme, ride in the display list as a `shadow` field, and paint on the canvas in device pixels. The display-list contract version moves from 6 to 7; a consumer pinning 6 must update.
