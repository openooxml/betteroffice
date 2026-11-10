---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
"@betteroffice/rust-crates": patch
---

Text in Word documents no longer forms ligatures such as fi and ti unless the document asks for them, matching Word. Contextual alternates also follow the document's formatting. Inside a formed ligature, the caret and selection now land between its letters.
