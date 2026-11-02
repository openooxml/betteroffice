---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
---

Text runs drawn without shaping carry the East Asian language of their Han text (`eastAsiaLanguage`), so the native renderer requests the Japanese, Korean or Traditional Chinese face the browser path already uses.
