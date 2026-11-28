---
"@betteroffice/pptx": minor
"@betteroffice/pptx-native": minor
"@betteroffice/pptx-i18n": minor
"@betteroffice/pptx-react": minor
"@betteroffice/rust-crates": minor
---

Adds opt-in `instantiatePlaceholders` to `insertSlide`, exposes `layouts` as `PresentationSlideLayoutSummary` in the worker command open result, and adds `placeholderType` to `TextBoxPrimitive`. In `PptxEditor`, inserting a slide now instantiates layout placeholders by default; untouched ones follow their layout and show its prompt or a localized guide.
