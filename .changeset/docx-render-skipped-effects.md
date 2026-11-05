---
'@betteroffice/rust-crates': minor
'@betteroffice/python-docx': minor
---

DOCX `render_png` no longer refuses a page over an effect it cannot draw: it draws text shadows, small caps and shape shadows and glows as the browser does, leaves other effects out, and reports them by kind in `skipped_effects`.
