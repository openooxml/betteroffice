# @betteroffice/docx-react

## 0.0.4

### Patch Changes

- 5c9a482: ArrowUp/ArrowDown move the caret by visual line with persistent goal-X (including across paragraphs, pages, columns, and into tables), and content below tables is clickable and editable again.
- 5c9a482: Collaborative presence: remote collaborator carets and selections render as colored overlays with name flags, anchored by yrs sticky indices so they rebase exactly under concurrent edits; carets follow remote typing instantly by inferring position from document updates.
- 5c9a482: Opening a document now seeds the collaborative session directly in the Rust engine instead of materializing the full TypeScript document model and projecting it; the TS model is built lazily only where the public API still exposes it, and the internal DrawingML host package is dissolved.
- 5c9a482: Remote collaborators' edits no longer move the local viewport: relayouts triggered by remote updates anchor to the topmost visible line via yrs sticky positions and compensate the scroll offset, while caret scrolling fires only for local actions. Anchoring holds across page boundaries too, so text overflowing onto a new page (or pulling back off one) no longer jumps the viewport for either the author or a viewer.
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
  - @betteroffice/docx@0.0.4
  - @betteroffice/docx-i18n@0.0.4

## 0.0.3

### Patch Changes

- Updated dependencies [b34bb01]
  - @betteroffice/docx@0.0.3
  - @betteroffice/docx-i18n@0.0.3

## 0.0.2

### Patch Changes

- eed05a6: Fix the published dependency ranges: 0.0.1 shipped the unresolved `workspace:*` protocol for `@betteroffice/docx` and `@betteroffice/docx-i18n`, which made `npm install @betteroffice/docx-react` fail. Ranges are now pinned to concrete versions at publish time.
  - @betteroffice/docx@0.0.2
  - @betteroffice/docx-i18n@0.0.2
