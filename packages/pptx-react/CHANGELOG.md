# @betteroffice/pptx-react

## 0.0.3

### Patch Changes

- 5212690: Google Slides-style editor toolbar for the PPTX editor: new-slide split button
  with layout picker, undo/redo, zoom, select and text-box tools, and contextual
  text formatting that also applies to whole shapes on selection. Text formatting
  now spans paragraph boundaries as a single undoable operation, double/triple
  click select word/paragraph, and roundRect corners render circular per the
  OOXML adj value instead of stretching with the shape.
- Updated dependencies [5212690]
  - @betteroffice/pptx@0.0.3
  - @betteroffice/pptx-i18n@0.0.3

## 0.0.2

### Patch Changes

- 64e5940: Add pointer-based shape movement and text range selection to the PPTX editor.
- 69d62f1: Refine the XLSX and PPTX editor toolbars with compact DOCX-style control rails,
  grouped icon actions, and responsive value fields.
  - @betteroffice/pptx@0.0.2
  - @betteroffice/pptx-i18n@0.0.2
