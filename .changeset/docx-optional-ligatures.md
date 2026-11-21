---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
"@betteroffice/rust-crates": minor
---

Ligatures such as fi require document opt-in, matching Word. Contextual alternates follow the document too. The caret and selection land between a ligature's letters, and selecting a field result covers all of it. Set new fields in `TextFormatting`, `RunFormatting`, `DocAttrs` and `DisplayTextContractMetadata` literals.
