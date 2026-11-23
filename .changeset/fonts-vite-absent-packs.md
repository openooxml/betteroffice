---
"@betteroffice/fonts": patch
---

An absent `@betteroffice/fonts-intl` or `@betteroffice/fonts-cjk` now counts as absent under Vite 7 and 8, as it does under esbuild and webpack, so documents that need those faces open with substitutes instead of failing.
