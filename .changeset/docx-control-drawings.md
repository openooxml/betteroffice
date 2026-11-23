---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": minor
---

Reads now include `drawings` on inline content-control atoms by default, including text boxes and revision views. Drawing geometry targets accept an optional `index` to address a drawing inside a control, including footer controls. Rust `TextAtom` literals must set `drawings`.
