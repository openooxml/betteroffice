---
"@betteroffice/docx": patch
---

With a bundled font source configured through `configureDefaultFonts`, document fonts load from that bundle and no request goes to Google Fonts. A family resolves by its Word name or by its bundled name, such as Gelasio or Noto Sans SC. A stalled load gives up after 5 seconds and is retried on the next request. Styles a family does not ship are left for the browser to synthesize, and faces that arrive after text was drawn repaint it.
