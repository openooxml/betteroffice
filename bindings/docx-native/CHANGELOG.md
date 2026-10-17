# @betteroffice/docx-native

## 0.5.0

### Patch Changes

- a90a6cd: Adds `@betteroffice/docx-native`, `@betteroffice/pptx-native` and `@betteroffice/xlsx-native` with prebuilt Rust engines for Node.js, Bun and Deno. The packages open, edit, render PNGs and save documents.
- c6ab4fa: Native page and slide rendering holds each image's decoder to the image budget left, so a GIF whose frame is larger than its declared size is skipped instead of allocating past the budget.
