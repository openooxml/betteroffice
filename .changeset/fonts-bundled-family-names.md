---
"@betteroffice/fonts": minor
---

`resolveBundledFamilyFace` and `resolveFamily` also resolve a family by its bundled name (e.g. `Gelasio`), returning only faces of the requested weight and style. **BREAKING:** custom `BundledFontSource` implementations must add `resolveFamily`; the bundled providers already do.
