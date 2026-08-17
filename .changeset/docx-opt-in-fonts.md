---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
"@betteroffice/fonts": patch
---

`@betteroffice/docx` no longer imports `@betteroffice/fonts` on its own, so installing the package does nothing by itself and its specifier is gone from the published bundle — an esbuild consumer without the optional peer builds again. Opt in with `configureDefaultFonts({ fonts })` after importing the module, or lazily with `configureDefaultFonts({ load: () => import('@betteroffice/fonts') })`; add `baseUrl` to either to serve the face binaries from a CDN.
