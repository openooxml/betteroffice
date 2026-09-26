---
"@betteroffice/pptx-react": minor
"@betteroffice/pptx": minor
"@betteroffice/pptx-i18n": minor
"@betteroffice/rust-crates": minor
---

Expose the PPTX editor's commands as `PptxEditorApi.commands`: serializable state with a stated reason for every disabled command, descriptors with platform-aware shortcuts, and `execute` that runs after pending input such as a decoding picture and checks availability again. Hosts compose built-in controls with their own actions through the `toolbar` and `showToolbar` props, `PptxCommandProvider`, `usePptxCommand`/`usePptxCommandState`, `EditorToolbar mode="commands"`, `ToolbarCommand`, `ToolbarCommandButton`, `ToolbarCommandSelect` and `ToolbarOverflow`, inside or outside the editor. The prop-based `EditorToolbar`, `Toolbar` and `useEditorToolbar` keep working and are deprecated. Canvas and panel proposal review, keyboard shortcuts and the default toolbar run through the same commands; a picture lands on the slide it was chosen for. Narrow toolbars move groups into a keyboard-accessible More menu, and `ToolbarDropdown` follows the menu or dialog pattern of its content. New locale keys cover command labels and disabled reasons. The command and toolbar composition API is experimental and may change in minor releases. `@betteroffice/pptx` adds `anchorCaret` and `resolveCaretAnchor`: plain-data caret positions that follow later edits, undo, redo and remote updates.
