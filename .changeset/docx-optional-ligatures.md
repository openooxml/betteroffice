---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
"@betteroffice/rust-crates": minor
---

Word documents no longer form ligatures such as fi unless the document asks for them, matching Word; contextual alternates follow the document too. The caret and selection land between a ligature's letters, and selecting a field result covers all of it. Set `ligatures` and `contextual_alternates` on `TextFormatting` and `RunFormatting` literals, and `font_features` on `DocAttrs` and `DisplayTextContractMetadata` literals.
