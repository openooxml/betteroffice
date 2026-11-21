---
'@betteroffice/python-docx': patch
'@betteroffice/python-pptx': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

The Python packages now install the bundled fonts as `betteroffice-fonts-latin`, `betteroffice-fonts-intl` and `betteroffice-fonts-cjk` and render with them from disk; nothing is downloaded. `register_font` still adds or replaces a family.
