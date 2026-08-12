---
'@betteroffice/fonts': patch
'@betteroffice/fonts-cjk': patch
'@betteroffice/docx': patch
'@betteroffice/docx-react': patch
---

Ship the bundled fonts as published packages and load them by default.

`@betteroffice/docx-fonts` was private and 404 on npm, so every consumer measured documents with synthetic metrics unless they hand-wired a provider. Across 813 real-world documents scored against Word's own `docProps/app.xml <Pages>`, that costs 15.4 points of exact page-count accuracy — 61.9% with real font metrics against 46.5% without, and a mean absolute error of 0.80 pages against 2.47.

The package is now `@betteroffice/fonts` and is publishable. It is shared rather than docx-specific — the PPTX demo already imported it — so the docx-prefixed name is gone. The five Noto CJK faces are 33 MB against 7.9 MB for everything else, so they move to a separate optional `@betteroffice/fonts-cjk`; npm has no partial-tarball fetch, so only a package boundary keeps those bytes out of installs that do not need them. Face resolution policy stays in one place, in `@betteroffice/fonts`.

`@betteroffice/docx` now resolves a default provider itself, so framework-agnostic and server-side consumers get real font metrics too rather than only React hosts. It resolves through an optional dynamic `import()` and is declared an optional peer dependency, so core never takes a static edge onto the font bytes and consumers who inject their own provider or omit the package pay nothing. An explicitly injected `measurementFontProvider` still wins and short-circuits the default entirely.

Faces load same-origin by default; `configureDefaultFonts({ baseUrl })` opts into a CDN. Same-origin stays the default deliberately — a CDN default would leak document-font usage to a third party and break offline and strict-CSP deployments.

When no provider resolves at all and a family falls through to synthetic metrics, the registry now warns once with what to install, instead of silently paginating wrong.
