---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/docx-i18n": patch
---

Suggesting mode now tracks character formatting and hyperlink edits through optional `suggesting` authors on `toggleMark`, `formatRange`, `clearFormatting` and `setHyperlink`. Character formatting over fields or inline content controls is refused with `SuggestedFormattingOverContainerError`; commands return `unsupported-suggesting`.
