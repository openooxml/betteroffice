# @betteroffice/vsdx-i18n

## 0.0.3

### Patch Changes

- 16d33a7: Fix locale declarations for TypeScript consumers with `skipLibCheck: false` and update React editors to depend on the corrected i18n packages.

## 0.0.2

### Patch Changes

- 0df1345: Make the editor canvas focusable and give it a keyboard layer: undo and redo, Delete or Backspace,
  arrow-key nudge with a larger Shift step, and Escape to cancel a gesture or clear the selection.
  Keys are ignored while focus is in a text input, and a handle resize of a locked shape is refused
  with a translated message.
- fa8e835: Replace the desktop-style ribbon with Visio for the web's flat command bar: a single 45px row of
  icon buttons separated by thin rules, with no group labels, under a tab strip whose active tab
  carries a text-width underline. Tabs with no commands say so instead of rendering blank.
- 3c20209: Add a right-click menu for the selected shape with keyboard navigation and z-order submenus that
  stay inside the viewport, draw the rotation grip as a circle on a stalk, and give the shapes panel
  a header, a search box and a category rail. A queued drag preview follows the latest Shift state.
- eda21ff: Add in-place shape text editing to the VSDX editor. Text edits run through the mutation policy with a typed receipt, are undoable and authorized on remote updates, and saving patches only the edited shape's `Text` element.
- 1391c9b: Expand the standard shape gallery to thirty-two shapes, including a real ellipse at a 3:2 aspect
  ratio. Every gallery preview is derived from the geometry the insert writes, so a shape lands in
  the proportions it was picked in.
