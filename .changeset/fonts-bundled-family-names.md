---
"@betteroffice/fonts": patch
---

`resolveBundledFamilyFace` and the provider's `resolveFamily` resolve a family by its bundled name (for example `Gelasio`, `Source Sans 3` or `Noto Sans SC`) as well as by its Word name. They return only a face whose weight and style match, so a browser can synthesize the styles a family doesn't ship.
