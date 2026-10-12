import { useCallback, useState, type MouseEvent } from 'react';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { TextContextMenu, type TextContextAction, type TextContextMenuItem } from '../../TextContextMenu';
import type { DocxCommandEnvironment } from '../../../commands/evaluate';
import type { DocxCommandStore } from '../../../commands/types';
import { normalizeInputEvent } from './inputEvents';
import { admitInputIntents } from './inputAdmission';

export function useWorkerContextMenu(options: {
  binding(): PeerlessEditSessionBinding | null;
  environment(): DocxCommandEnvironment;
  commands: DocxCommandStore;
  clipboard(cut: boolean): Promise<void>;
  onPaste?: () => void;
  onError?: (error: Error) => void;
}) {
  const [menu, setMenu] = useState<{
    binding: PeerlessEditSessionBinding; snapshot: ReturnType<PeerlessEditSessionBinding['viewStore']['getSnapshot']>;
    position: { x: number; y: number };
  } | null>(null);
  const close = useCallback(() => setMenu(null), []);
  const open = (event: MouseEvent) => {
    const binding = options.binding();
    if (!binding) return;
    event.preventDefault(); event.stopPropagation();
    setMenu({ binding, snapshot: structuredClone(binding.viewStore.getSnapshot()), position: { x: event.clientX, y: event.clientY } });
  };
  const selection = menu?.snapshot.workerSelection;
  const hasSelection = !!selection && (selection.anchor.story !== selection.focus.story ||
    selection.anchor.paraId !== selection.focus.paraId || selection.anchor.offset !== selection.focus.offset);
  const readOnly = options.environment().readOnly || options.environment().mode === 'viewing';
  const t = options.environment().translate;
  const items: TextContextMenuItem[] = [
    { action: 'copy', label: t('contextMenu.copy'), disabled: !hasSelection },
    ...readOnly ? [] : [
      { action: 'cut' as const, label: t('contextMenu.cut'), disabled: !hasSelection },
      { action: 'paste' as const, label: t('contextMenu.paste') },
      { action: 'delete' as const, label: t('contextMenu.delete'), disabled: !hasSelection },
    ],
    { action: 'selectAll', label: t('contextMenu.selectAll') },
  ];
  if (menu?.snapshot.editState?.tableContext && !readOnly) {
    for (const [action, key] of [['addRowAbove', 'insertRowAbove'], ['addRowBelow', 'insertRowBelow'], ['deleteRow', 'deleteRow'],
      ['addColumnLeft', 'insertColumnLeft'], ['addColumnRight', 'insertColumnRight'], ['deleteColumn', 'deleteColumn'],
      ['mergeCells', 'mergeCells'], ['splitCell', 'splitCell'], ['deleteTable', 'deleteTable']] as const) {
      items.push({ action, label: t(`table.${key}`), disabled: !options.commands.getState('tableAction', action).enabled });
    }
  }
  const action = async (action: TextContextAction) => {
    if (!menu) return;
    close();
    try {
      const current = options.binding();
      if (current !== menu.binding) throw new Error('Document was replaced');
      const snapshot = current.viewStore.getSnapshot();
      if (snapshot.version !== menu.snapshot.version || JSON.stringify(snapshot.workerSelection) !== JSON.stringify(selection)) {
        throw new Error('Context menu target changed');
      }
      if (action === 'copy' || action === 'cut') { await options.clipboard(action === 'cut'); return; }
      if (action === 'selectAll') {
        current.commands.submit({ kind: 'navigate', payload: { unit: 'document-boundary', direction: 'backward', extend: false } }, { source: 'selection', fence: true });
        current.commands.submit({ kind: 'navigate', payload: { unit: 'document-boundary', direction: 'forward', extend: true } }, { source: 'selection', fence: true });
        return;
      }
      if (options.environment().readOnly || options.environment().mode === 'viewing') return;
      if (action === 'delete' && selection) current.commands.submit({ kind: 'replaceSelection', payload: { text: '', selection } }, { fence: true, source: 'paste' });
      else if (action === 'paste' || action === 'pasteAsPlainText') {
        const text = await navigator.clipboard.readText();
        if (current !== options.binding()) throw new Error('Document was replaced');
        const latest = current.viewStore.getSnapshot();
        if (latest.version !== menu.snapshot.version || JSON.stringify(latest.workerSelection) !== JSON.stringify(selection)) {
          throw new Error('Context menu target changed');
        }
        if (options.environment().readOnly || options.environment().mode === 'viewing') return;
        const intents = normalizeInputEvent({ source: 'paste', text, eventId: 1, inputTime: performance.now() },
          { composing: false, compositionId: null, lastHandledEventId: null, trailingCompositionText: null }).intents;
        if (selection && intents[0]?.kind === 'replaceSelection') intents[0].payload.selection = selection;
        admitInputIntents(current.admission, intents, { source: 'paste', fence: true });
        options.onPaste?.();
      } else await options.commands.execute('tableAction', action as never);
    } catch (error) { options.onError?.(error instanceof Error ? error : new Error(String(error))); }
  };
  return { open, element: menu && <TextContextMenu isOpen position={menu.position} hasSelection={hasSelection}
    isEditable={!readOnly} items={items} onAction={(value) => { void action(value); }} onClose={close} /> };
}
