---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
"@betteroffice/rust-crates": patch
---

Text boxes apply all package paragraph and character styles by default, including inherited fonts, colors, languages and spacing. Shared state preserves these styles across peers and payload edits without changing saved payloads. `font_preflight` includes styled text-box faces and inherited language scripts.
