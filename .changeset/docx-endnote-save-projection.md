---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Edits made in an endnote now reach the saved file. Endnote stories were seeded into the collaborative document as `en:{id}` alongside the footnotes' `fn:{id}`, but the save projection never read them back: it collected base stories for the body, header/footer parts and footnotes only, and rebuilt just those parts of the package. An endnote story was therefore write-only — every edit in it was silently replaced by the imported text on save.

Both note families now project through one path, so a footnote and an endnote round-trip the same way, and a projected note drops its `verbatimXml` so the modeled content is what gets written rather than the original XML. The editor's dirty-story tracking follows: direct input in an `en:` story marks that story for projection instead of falling back to the body, which is what gated the partial (changed-stories-only) save.
