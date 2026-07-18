import { useEffect } from 'react';
import type { useTableSelection } from '../../../hooks/useTableSelection';
import type { useFindReplace } from '../../../hooks/useFindReplace';
import type { useHyperlinkDialog } from '../../dialogs/HyperlinkDialog';
import type { PagedEditorRef } from '../PagedEditor';
import { yrsHyperlinkAtSelection, yrsSelectedText } from '../yrsCommands';

/**
 * Top-level keyboard shortcuts:
 *  - Cmd/Ctrl+O → open the DOCX picker when File > Open is enabled
 *  - Cmd/Ctrl+F → open Find dialog (seeded with current selection)
 *  - Cmd/Ctrl+H → open Find/Replace dialog
 *  - Cmd/Ctrl+K → open Hyperlink dialog (edit if cursor sits on a link)
 *  - Delete/Backspace on a full-table layout selection → delete the table
 *
 * Listens on `document` so the shortcut works even when focus isn't in the
 * editor. `disableFindReplaceShortcuts` lets the host app reclaim Cmd+F /
 * Cmd+H when the editor is embedded inside another shell.
 */
export function useKeyboardShortcuts({
  pagedEditorRef,
  disableFindReplaceShortcuts,
  showFileOpen,
  onOpenDocument,
  findReplace,
  hyperlinkDialog,
  tableSelection,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  disableFindReplaceShortcuts: boolean;
  showFileOpen: boolean;
  onOpenDocument?: () => void;
  findReplace: ReturnType<typeof useFindReplace>;
  hyperlinkDialog: ReturnType<typeof useHyperlinkDialog>;
  tableSelection: ReturnType<typeof useTableSelection>;
}) {
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const isMac = navigator.platform.toUpperCase().indexOf('MAC') >= 0;
      const cmdOrCtrl = isMac ? e.metaKey : e.ctrlKey;

      // Delete a layout-selected table.
      if (!cmdOrCtrl && !e.shiftKey && !e.altKey) {
        if (e.key === 'Delete' || e.key === 'Backspace') {
          if (tableSelection.state.tableIndex !== null) {
            e.preventDefault();
            tableSelection.handleAction('deleteTable');
            return;
          }
        }
      }

      if (cmdOrCtrl && !e.shiftKey && !e.altKey) {
        if (e.key.toLowerCase() === 'f') {
          if (disableFindReplaceShortcuts) return;
          e.preventDefault();
          const selection = window.getSelection();
          const selectedText = selection && !selection.isCollapsed ? selection.toString() : '';
          findReplace.openFind(selectedText);
        } else if (e.key.toLowerCase() === 'o') {
          if (!showFileOpen || !onOpenDocument) return;
          e.preventDefault();
          onOpenDocument();
        } else if (e.key.toLowerCase() === 'h') {
          if (disableFindReplaceShortcuts) return;
          e.preventDefault();
          const selection = window.getSelection();
          const selectedText = selection && !selection.isCollapsed ? selection.toString() : '';
          findReplace.openReplace(selectedText);
        } else if (e.key.toLowerCase() === 'k') {
          e.preventDefault();
          const session = pagedEditorRef.current?.getYrsSession();
          if (session) {
            const selectedText = yrsSelectedText(session);
            const existingLink = yrsHyperlinkAtSelection(session);
            if (existingLink) {
              hyperlinkDialog.openEdit({
                url: existingLink.href,
                displayText: selectedText || existingLink.text,
                tooltip: existingLink.tooltip,
              });
            } else {
              hyperlinkDialog.openInsert(selectedText);
            }
          }
        }
      }
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [
    pagedEditorRef,
    disableFindReplaceShortcuts,
    showFileOpen,
    onOpenDocument,
    findReplace,
    hyperlinkDialog,
    tableSelection,
  ]);
}
