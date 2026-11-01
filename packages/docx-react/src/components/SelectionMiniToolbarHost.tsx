import { forwardRef, useCallback } from 'react';
import { useTranslation } from '../i18n';
import { useDocxCommandState, useDocxCommands } from '../commands/hooks';
import { commandShortcut } from '../commands/descriptors';
import { formatKeys } from './dialogs/KeyboardShortcutsDialog/ShortcutItem';
import { restoreEditorFocus } from './toolbar/activation';
import type { DocxCommandId } from '../commands/types';
import type { DocxSelectionToolbarItem } from './DocxEditor';
import {
  SelectionMiniToolbar,
  type SelectionMiniToolbarButton,
  type SelectionToolbarAnchorRect,
} from './SelectionMiniToolbar';

const ITEM_ICONS: Record<DocxSelectionToolbarItem, string> = {
  clearFormatting: 'format_clear',
  pastePlainText: 'content_paste',
  bold: 'format_bold',
  italic: 'format_italic',
  underline: 'format_underlined',
  strikethrough: 'strikethrough_s',
};

const COMMAND_ITEMS: ReadonlySet<DocxSelectionToolbarItem> = new Set([
  'clearFormatting',
  'bold',
  'italic',
  'underline',
  'strikethrough',
]);

export interface SelectionMiniToolbarHostProps {
  anchorRect: SelectionToolbarAnchorRect;
  items: readonly DocxSelectionToolbarItem[];
  /** Whether Paste as plain text is available in this editor arm. */
  pasteEnabled: boolean;
  onPastePlainText: () => void;
  onDismiss: (returnFocus: boolean) => void;
}

/**
 * Binds the selection mini toolbar to the editor command store. An item
 * whose command is unavailable (for example in a read-only editor) is
 * not shown.
 */
export const SelectionMiniToolbarHost = forwardRef<HTMLDivElement, SelectionMiniToolbarHostProps>(
  function SelectionMiniToolbarHost({ anchorRect, items, pasteEnabled, onPastePlainText, onDismiss }, ref) {
    const { t } = useTranslation();
    const store = useDocxCommands();
    const clearFormatting = useDocxCommandState('clearFormatting');
    const bold = useDocxCommandState('bold');
    const italic = useDocxCommandState('italic');
    const underline = useDocxCommandState('underline');
    const strikethrough = useDocxCommandState('strikethrough');
    const states = { clearFormatting, bold, italic, underline, strikethrough };

    const dismiss = useCallback(
      (returnFocus: boolean) => {
        onDismiss(returnFocus);
        if (returnFocus) restoreEditorFocus(store);
      },
      [onDismiss, store]
    );

    const buttons: SelectionMiniToolbarButton[] = [];
    for (const item of items) {
      if (item === 'pastePlainText') {
        if (!pasteEnabled) continue;
        buttons.push({
          id: item,
          label: t('contextMenu.pastePlainText'),
          shortcut: formatKeys(t('contextMenu.pastePlainTextShortcut')),
          icon: ITEM_ICONS[item],
          onActivate: () => {
            onDismiss(false);
            restoreEditorFocus(store);
            onPastePlainText();
          },
        });
        continue;
      }
      if (!COMMAND_ITEMS.has(item)) continue;
      const state = states[item as keyof typeof states];
      if (!state.enabled) continue;
      const id = item as DocxCommandId;
      buttons.push({
        id: item,
        label: t(store.getDescriptor(id).labelKey),
        shortcut: commandShortcut(id),
        icon: ITEM_ICONS[item],
        active: state.active === true,
        onActivate: () => {
          onDismiss(false);
          void store.execute(id, null as never).then(() => restoreEditorFocus(store));
        },
      });
    }

    if (buttons.length === 0) return null;
    return (
      <SelectionMiniToolbar
        ref={ref}
        anchorRect={anchorRect}
        buttons={buttons}
        label={t('selectionToolbar.label')}
        hint={t('selectionToolbar.focusHint')}
        onDismiss={dismiss}
      />
    );
  }
);

export type { SelectionToolbarAnchorRect };
