import { useCallback } from 'react';
import type { FormattingAction } from '../../Toolbar';
import type { useHyperlinkDialog } from '../../dialogs/HyperlinkDialog';
import type { PagedEditorRef } from '../PagedEditor';

/** Toolbar and structural commands routed to the authoritative Yrs editor. */
export function useFormattingActions({
  focusActiveEditor,
  pagedEditorRef,
  hyperlinkDialog,
}: {
  focusActiveEditor: () => void;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  hyperlinkDialog: ReturnType<typeof useHyperlinkDialog>;
}) {
  const handleFormat = useCallback(
    (action: FormattingAction) => {
      if (action === 'insertLink') {
        hyperlinkDialog.openInsert('');
        return;
      }
      pagedEditorRef.current?.applyYrsFormatting(action);
    },
    [hyperlinkDialog, pagedEditorRef]
  );

  const applyCommand = useCallback(
    (command: Parameters<PagedEditorRef['applyYrsCommand']>[0]) => {
      pagedEditorRef.current?.applyYrsCommand(command);
      focusActiveEditor();
    },
    [focusActiveEditor, pagedEditorRef]
  );

  const handleInsertTable = useCallback(
    (rows: number, columns: number) => applyCommand({ type: 'insertTable', rows, columns }),
    [applyCommand]
  );
  const handleInsertPageBreak = useCallback(
    () => applyCommand({ type: 'insertPageBreak' }),
    [applyCommand]
  );
  const handleInsertSectionBreakNextPage = useCallback(
    () => applyCommand({ type: 'insertSectionBreak', breakType: 'nextPage' }),
    [applyCommand]
  );
  const handleInsertSectionBreakContinuous = useCallback(
    () => applyCommand({ type: 'insertSectionBreak', breakType: 'continuous' }),
    [applyCommand]
  );

  return {
    handleFormat,
    handleInsertTable,
    handleInsertPageBreak,
    handleInsertSectionBreakNextPage,
    handleInsertSectionBreakContinuous,
    // TOC insertion has no Yrs command yet; keep the UI callback inert.
    handleInsertTOC: useCallback(() => focusActiveEditor(), [focusActiveEditor]),
  };
}
