import { useCallback, useState } from 'react';
import { toast } from 'sonner';
import type { HyperlinkData, useHyperlinkDialog } from '../../dialogs/HyperlinkDialog';
import type { HyperlinkPopupData } from '../../ui/HyperlinkPopup';
import type { YrsEditorCommand } from '../yrsCommands';

/** Where a hyperlink edit was opened: the link dialog or the link popup. */
export type HyperlinkSource = 'link' | 'linkPopup';

/**
 * Owns the dialog-driven hyperlink flow (insert / edit / remove) and the
 * Google-Docs-style floating popup that opens when the cursor lands on
 * an existing link. The dialog handle (`hyperlinkDialog`) is owned by
 * the parent and threaded in; the `insertLink` command opens it.
 * `applyCommand` writes against the selection its source opened with.
 */
export function useHyperlinkActions({
  hyperlinkDialog,
  openPopup,
  applyCommand,
  focusActiveEditor,
}: {
  hyperlinkDialog: ReturnType<typeof useHyperlinkDialog>;
  openPopup: () => void;
  applyCommand: (source: HyperlinkSource, command: YrsEditorCommand) => void;
  focusActiveEditor: () => void;
}) {
  const [hyperlinkPopupData, setHyperlinkPopupData] = useState<HyperlinkPopupData | null>(null);

  const handleHyperlinkSubmit = useCallback(
    (data: HyperlinkData) => {
      applyCommand('link', {
        type: 'setHyperlink',
        href: data.url || '',
        tooltip: data.tooltip,
        displayText: data.displayText,
        editExisting: hyperlinkDialog.state.isEditing,
      });
      hyperlinkDialog.close();
      focusActiveEditor();
    },
    [applyCommand, hyperlinkDialog, focusActiveEditor]
  );

  const doRemoveHyperlink = useCallback(() => {
    applyCommand('link', { type: 'removeHyperlink' });
    focusActiveEditor();
  }, [applyCommand, focusActiveEditor]);

  const handleHyperlinkRemove = useCallback(() => {
    doRemoveHyperlink();
    hyperlinkDialog.close();
  }, [hyperlinkDialog, doRemoveHyperlink]);

  const handleHyperlinkClick = useCallback(
    (data: HyperlinkPopupData) => {
      openPopup();
      setHyperlinkPopupData(data);
    },
    [openPopup]
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
      applyCommand('linkPopup', {
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
    [applyCommand, focusActiveEditor, hyperlinkPopupData]
  );

  const handleHyperlinkPopupRemove = useCallback(() => {
    applyCommand('linkPopup', {
      type: 'removeHyperlink',
      href: hyperlinkPopupData?.href,
    });
    setHyperlinkPopupData(null);
    focusActiveEditor();
    toast('Link removed');
  }, [applyCommand, focusActiveEditor, hyperlinkPopupData]);

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
