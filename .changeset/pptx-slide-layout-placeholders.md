---
"@betteroffice/pptx": minor
"@betteroffice/pptx-native": minor
"@betteroffice/pptx-i18n": minor
"@betteroffice/pptx-react": minor
"@betteroffice/rust-crates": minor
---

Adds opt-in `instantiatePlaceholders` to `insertSlide`, exposes `layouts` and `PresentationSlideLayoutSummary` in worker sessions, and adds `placeholderType` to `TextBoxPrimitive`. In `PptxEditor`, inserting a slide now instantiates layout placeholders by default; untouched placeholders follow their layout, and empty ones show the layout prompt or a localized guide.
