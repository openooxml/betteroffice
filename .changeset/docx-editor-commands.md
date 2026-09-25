---
"@betteroffice/docx-react": minor
"@betteroffice/docx": minor
"@betteroffice/docx-i18n": minor
"@betteroffice/rust-crates": minor
---

Expose the DOCX editor's commands as `DocxEditorRef.commands`: serializable state with a stated reason for every disabled command, descriptors with shortcuts, and `execute` that runs after accepted input and checks availability again. Hosts compose built-in controls with their own actions through the `toolbar` prop, `DocxCommandProvider`, `useDocxCommand`/`useDocxCommandState`, `EditorToolbar` (now including `EditorToolbar.Review`), `ToolbarCommand`, `ToolbarCommandButton`, `ToolbarCommandSelect`, `ToolbarButton`, `ToolbarGroup`, `ToolbarSeparator` and `ToolbarOverflow`, inside or outside the editor. Dialogs and pickers a command opens apply to the document and selection they opened with (failing with `document-replaced` or `target-changed` otherwise), replacements and print wait for accepted input, and engine refusals fail the command. Narrow toolbars move groups into a keyboard-accessible More menu that keeps every choice of the built-in controls, and the keyboard help lists the command bindings. The selection context reports superscript, subscript and highlight, and `toggleMark` accepts superscript and subscript. New locale keys cover command labels and disabled reasons.
