---
"@betteroffice/fonts": patch
---

`resolveBundledFamilyFace` and `resolveFamily` now resolve a font family by its bundled name (e.g. `Gelasio`) as well as its Word name, returning only faces matching the requested weight and style.
