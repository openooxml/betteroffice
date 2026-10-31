---
"@betteroffice/rust-crates": minor
---

The docx, pptx, xlsx and vsdx crates now resolve fonts by default: document-embedded faces, then faces you register, then the bundled Latin, international and CJK packs, read locally. Use `open_with_options` to pick a source, or turn off default features.
