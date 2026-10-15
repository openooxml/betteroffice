import { useCallback, useState, type MouseEvent } from 'react';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { TextContextMenu, type TextContextAction, type TextContextMenuItem } from '../../TextContextMenu';
import type { DocxCommandEnvironment } from '../../../commands/evaluate';
import type { DocxCommandStore } from '../../../commands/types';
import { normalizeInputEvent } from './inputEvents';
import { admitInputIntents } from './inputAdmission';
import { isPasteFragmentAllowed } from './clipboard';
import { parseClipboardHtml } from '../internals/clipboardHtml';
import { WorkerObjectDialog } from './WorkerObjectDialog';
import type { PeerlessPropertyIntent } from './inputEvents';
import type { DocxCommandResult } from '../../../commands/types';
import type { DocxClipboardFragment } from '@betteroffice/docx/yrs';

async function readClipboardFragment(): Promise<DocxClipboardFragment | null> {
  if (typeof navigator === 'undefined' || !navigator.clipboard || !('read' in navigator.clipboard)) return null;
  let items: ClipboardItem[];
  try {
    items = await navigator.clipboard.read();
  } catch {
    return null;
  }
  for (const item of items) {
    if (!item.types.includes('text/html')) continue;
    let html: string;
    try {
      html = await (await item.getType('text/html')).text();
    } catch {
      return null;
    }
    return parseClipboardHtml(html);
  }
  return null;
}

export function useWorkerContextMenu(options: {
  binding(): PeerlessEditSessionBinding | null;
  environment(): DocxCommandEnvironment;
  commands: DocxCommandStore;
  clipboard(cut: boolean): Promise<void>;
  submit?(intent: PeerlessPropertyIntent): Promise<DocxCommandResult>;
  onPaste?: () => void;
  onError?: (error: Error) => void;
}) {
  const [menu, setMenu] = useState<{
    binding: PeerlessEditSessionBinding; snapshot: ReturnType<PeerlessEditSessionBinding['viewStore']['getSnapshot']>;
    position: { x: number; y: number };
    object?: { kind: 'image' | 'cell'; handle: string; version: string; x: number; y: number; width: number; height: number };
  } | null>(null);
  const [dialog, setDialog] = useState<{ menu: NonNullable<typeof menu>; action: 'resizeImage' | 'moveImage' | 'resizeColumn' } | null>(null);
  const close = useCallback(() => setMenu(null), []);
  const open = (event: MouseEvent) => {
    const binding = options.binding();
    if (!binding) return;
    event.preventDefault(); event.stopPropagation();
    const snapshot = structuredClone(binding.viewStore.getSnapshot());
    const geometry = snapshot.geometry;
    const objects = geometry?.objects;
    const surface = event.target instanceof Element ? event.target.closest<HTMLElement>('.paged-editor__page') : null;
    const page = geometry?.pages.find(page => page.page.pageIndex === Number(surface?.dataset.pageIndex));
    const box = surface?.getBoundingClientRect();
    let object: NonNullable<typeof menu>['object'];
    if (objects && geometry && page && box?.width && box.height && objects.version === geometry.documentVersion &&
      objects.frameEpoch === geometry.frameEpoch && objects.layoutRevision === geometry.layoutRevision) {
      const x = (event.clientX - box.left) * page.bounds.width / box.width;
      const y = (event.clientY - box.top) * page.bounds.height / box.height;
      for (const kind of ['image', 'cell'] as const) {
        const found = objects.objects.find(target => target.kind === kind && target.rects.some(rect => rect.pageIndex === page.page.pageIndex &&
          x >= rect.x && x <= rect.x + rect.width && y >= rect.y && y <= rect.y + rect.height));
        const rect = found?.rects.find(rect => rect.pageIndex === page.page.pageIndex && x >= rect.x && x <= rect.x + rect.width && y >= rect.y && y <= rect.y + rect.height);
        if (found && rect) { object = { kind, handle: found.handle, version: objects.version, x: rect.x, y: rect.y, width: rect.width, height: rect.height }; break; }
      }
    }
    setMenu({ binding, snapshot, object, position: { x: event.clientX, y: event.clientY } });
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
  if (menu?.object && !readOnly && options.submit) {
    if (menu.object.kind === 'image') items.unshift({ action: 'resizeImage', label: 'Resize image' }, { action: 'moveImage', label: 'Move image', dividerAfter: true });
    else items.unshift({ action: 'resizeColumn', label: 'Column width', dividerAfter: true });
  }
  if (menu?.snapshot.editState?.tableContext && !readOnly) {
    for (const [action, key] of [['addRowAbove', 'insertRowAbove'], ['addRowBelow', 'insertRowBelow'], ['deleteRow', 'deleteRow'],
      ['addColumnLeft', 'insertColumnLeft'], ['addColumnRight', 'insertColumnRight'], ['deleteColumn', 'deleteColumn'],
      ['mergeCells', 'mergeCells'], ['splitCell', 'splitCell'], ['selectTable', 'selectTable'], ['deleteTable', 'deleteTable']] as const) {
      items.push({ action, label: t(`table.${key}`), disabled: !options.commands.getState('tableAction', action).enabled });
    }
  }
  const action = async (action: TextContextAction) => {
    if (!menu) return;
    if (action === 'resizeImage' || action === 'moveImage' || action === 'resizeColumn') {
      setDialog({ menu, action }); close(); return;
    }
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
        if (action === 'paste') {
          const fragment = await readClipboardFragment();
          if (fragment) {
            if (current !== options.binding()) throw new Error('Document was replaced');
            const latest = current.viewStore.getSnapshot();
            if (latest.version !== menu.snapshot.version || JSON.stringify(latest.workerSelection) !== JSON.stringify(selection)) {
              throw new Error('Context menu target changed');
            }
            if (options.environment().readOnly || options.environment().mode === 'viewing') return;
            if (isPasteFragmentAllowed(options.commands)) {
              admitInputIntents(current.admission, [{ kind: 'pasteFragment', payload: { fragment } }],
                { source: 'paste', fence: true });
              options.onPaste?.();
              return;
            }
          }
        }
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
  const applyObject = async (intent: PeerlessPropertyIntent) => {
    const origin = dialog?.menu;
    if (!origin || !options.submit) return false;
    try {
      const current = options.binding();
      const latest = current?.viewStore.getSnapshot();
      if (current !== origin.binding || latest?.version !== origin.snapshot.version ||
        JSON.stringify(latest.identity) !== JSON.stringify(origin.snapshot.identity)) throw new Error('Context menu target changed');
      if (options.environment().readOnly || options.environment().mode !== 'editing') throw new Error('Document is not editable');
      const result = await options.submit(intent);
      if (!result.ok) throw new Error(result.failure.message);
      return true;
    } catch (error) { options.onError?.(error instanceof Error ? error : new Error(String(error))); return false; }
  };
  return { open, element: <>{menu && <TextContextMenu isOpen position={menu.position} hasSelection={hasSelection}
    isEditable={!readOnly} items={items} onAction={(value) => { void action(value); }} onClose={close} />}
    {dialog?.menu.object && <WorkerObjectDialog action={dialog.action} target={dialog.menu.object} apply={applyObject} close={() => setDialog(null)} />}</> };
}
