# @betteroffice/pptx-native

## 0.6.0

### Patch Changes

- ad45f86: Adds opt-in `decodeFontFace` for packed fonts. Format package roots return a promise of `font` and optional `payload` ArrayBuffers; native bindings return byte arrays synchronously. Invalid options and decode errors reject or throw without retrying another decoder.
- ec04baa: The native packages require Node.js 20.16+ or 22.3+, which they need to read installed font packs from disk.
- 0530412: The same font bytes registered under several family names are now held once and no longer count against the 256-face limit. Opening a deck with more than 256 distinct faces now fails with `PptxResourceLimitError` instead of a closed worker connection.
- e664196: The PowerPoint engine is about 3 KB smaller compressed: the slide layout cache compares slide floats with a direct walk instead of a generic serializer. Layout and cache behaviour are unchanged.
- ec04baa: `@betteroffice/pptx-native` now installs `@betteroffice/fonts-latin`, `@betteroffice/fonts-intl` and `@betteroffice/fonts-cjk` and reads their faces from disk on demand, so Latin, international and CJK text renders with real faces without any fetch. On by default.
- db4ab57: Now exposes exactly the `@betteroffice/pptx` API (same entry points, exports and types) on the native engine for Node.js, Bun and Deno; pick the engine by import. The `browser` condition throws. The `openPresentation(Buffer)`/`Presentation` facade is removed.
- 8b61ea4: The PowerPoint engine can list the font faces a slide or deck asks for before layout, so hosts can load exactly those fonts before the first paint. Rendering is unchanged.
- 1398cff: Default fonts now load CJK script coverage for unmapped families and mixed text when the CJK packs are available; other scripts keep the browser font. `requestedFontFacesJson` adds optional `scripts` per face, and presentations with only coverage fonts render successfully.
- Updated dependencies [de8f092]
- Updated dependencies [5999962]
- Updated dependencies [9ab2b61]
- Updated dependencies [39ae19a]
- Updated dependencies [ec04baa]
- Updated dependencies [df33f6a]
- Updated dependencies [df33f6a]
- Updated dependencies [37199d9]
- Updated dependencies [842d624]
- Updated dependencies [e65df1c]
- Updated dependencies [ad45f86]
  - @betteroffice/fonts-cjk@1.0.0
  - @betteroffice/fonts@1.0.0
  - @betteroffice/fonts-intl@1.0.0
  - @betteroffice/fonts-latin@1.0.0

## 0.5.1

## 0.5.0

### Patch Changes

- c688743: Adds prebuilt Rust engines for Node.js, Bun and Deno that open, edit, render PNGs and save documents. Image decoding stays within the render budget.
