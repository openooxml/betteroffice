import { useCallback, useState } from 'react';
import { toast } from 'sonner';
import type { HyperlinkData, useHyperlinkDialog } from '../../dialogs/HyperlinkDialog';
import type { HyperlinkPopupData } from '../../ui/HyperlinkPopup';
import type { PagedEditorRef } from '../PagedEditor';

/**
 * Owns the dialog-driven hyperlink flow (insert / edit / remove) and the
 * Google-Docs-style floating popup that opens when the cursor lands on
 * an existing link. The dialog handle (`hyperlinkDialog`) is owned by
 * the parent and threaded in — Cmd/Ctrl+K in `useKeyboardShortcuts`
 * also opens it.
 */
export function useHyperlinkActions({
  hyperlinkDialog,
  pagedEditorRef,
  focusActiveEditor,
}: {
  hyperlinkDialog: ReturnType<typeof useHyperlinkDialog>;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  focusActiveEditor: () => void;
}) {
  const [hyperlinkPopupData, setHyperlinkPopupData] = useState<HyperlinkPopupData | null>(null);

  const handleHyperlinkSubmit = useCallback(
    (data: HyperlinkData) => {
      pagedEditorRef.current?.applyYrsCommand({
        type: 'setHyperlink',
        href: data.url || '',
        tooltip: data.tooltip,
        displayText: data.displayText,
        editExisting: hyperlinkDialog.state.isEditing,
      });
      hyperlinkDialog.close();
      focusActiveEditor();
    },
    [hyperlinkDialog, pagedEditorRef, focusActiveEditor]
  );

  const doRemoveHyperlink = useCallback(() => {
    pagedEditorRef.current?.applyYrsCommand({ type: 'removeHyperlink' });
    focusActiveEditor();
  }, [pagedEditorRef, focusActiveEditor]);

  const handleHyperlinkRemove = useCallback(() => {
    doRemoveHyperlink();
    hyperlinkDialog.close();
  }, [hyperlinkDialog, doRemoveHyperlink]);

  const handleHyperlinkClick = useCallback(
    (data: HyperlinkPopupData) => setHyperlinkPopupData(data),
    []
  );

  const handleHyperlinkPopupNavigate = useCallback((href: string) => {
    window.open(href, '_blank', 'noopener,noreferrer');
  }, []);

  const handleHyperlinkPopupCopy = useCallback((href: string) => {
    navigator.clipboard.writeText(href).catch(() => {
      // Fallback for browsers without async clipboard (older Safari, embedded webviews)
      const textarea = document.createElement('textarea');
      textarea.value = href;
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
    });
  }, []);

  const handleHyperlinkPopupEdit = useCallback(
    (displayText: string, href: string) => {
      pagedEditorRef.current?.applyYrsCommand({
        type: 'setHyperlink',
        href,
        tooltip: hyperlinkPopupData?.tooltip,
        displayText,
        editExisting: true,
        matchHref: hyperlinkPopupData?.href,
      });
      setHyperlinkPopupData(null);
      focusActiveEditor();
    },
    [focusActiveEditor, hyperlinkPopupData, pagedEditorRef]
  );

  const handleHyperlinkPopupRemove = useCallback(() => {
    pagedEditorRef.current?.applyYrsCommand({
      type: 'removeHyperlink',
      href: hyperlinkPopupData?.href,
    });
    setHyperlinkPopupData(null);
    focusActiveEditor();
    toast('Link removed');
  }, [focusActiveEditor, hyperlinkPopupData, pagedEditorRef]);

  const handleHyperlinkPopupClose = useCallback(() => {
    setHyperlinkPopupData(null);
  }, []);

  return {
    hyperlinkPopupData,
    setHyperlinkPopupData,
    handleHyperlinkSubmit,
    handleHyperlinkRemove,
    handleHyperlinkClick,
    handleHyperlinkPopupNavigate,
    handleHyperlinkPopupCopy,
    handleHyperlinkPopupEdit,
    handleHyperlinkPopupRemove,
    handleHyperlinkPopupClose,
  };
}
