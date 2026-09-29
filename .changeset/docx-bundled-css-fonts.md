---
"@betteroffice/docx": patch
---

With a bundled font source configured through `configureDefaultFonts`, document fonts load from that bundle and no request goes to Google Fonts. A family resolves by its Word name or by its bundled name, such as Gelasio or Noto Sans SC. When the bundle lacks a document font's CSS equivalent but carries the font itself, such as Comic Sans MS, the equivalent name gets the font's own bundled faces. A stalled or failed face gives up after 5 seconds and is retried on the next request, while the faces that did load are used. An older `@betteroffice/fonts` without `resolveFamily` registers only Regular faces. Styles a family does not ship are left for the browser to synthesize, and faces that arrive after text was drawn repaint it.
