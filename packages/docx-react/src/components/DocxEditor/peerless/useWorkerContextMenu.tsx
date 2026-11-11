import { useCallback, useState, type MouseEvent } from 'react';
import { IMAGE_LAYOUT_OPTIONS, type ImageLayoutTarget } from '@betteroffice/docx/docx';
import type { TranslationKey } from '@betteroffice/docx-i18n';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { TextContextMenu, type TextContextAction, type TextContextMenuItem } from '../../TextContextMenu';
import { evaluateDocxCommand, type DocxCommandEnvironment } from '../../../commands/evaluate';
import type { DocxCommandId, DocxCommandStore } from '../../../commands/types';
import { storeAriaShortcut, storeShortcutLabel } from '../../../commands/shortcuts';
import { normalizeInputEvent } from './inputEvents';
import { admitInputIntents, submitPasteFragment } from './inputAdmission';
import { isPasteFragmentAllowed } from './clipboard';
import { readClipboardFragment } from '../internals/clipboardHtml';
import { WorkerObjectDialog } from './WorkerObjectDialog';
import type { PeerlessPropertyIntent } from './inputEvents';
import type { DocxCommandResult } from '../../../commands/types';
import type { PeerlessViewSnapshot } from './viewStore';

const IMAGE_WRAP_TARGETS: ReadonlySet<string> = new Set<ImageLayoutTarget>([
  ...IMAGE_LAYOUT_OPTIONS.map(option => option.choice),
  'square',
  'tight',
  'through',
]);

function isImageWrapTarget(value: string): value is ImageLayoutTarget {
  return IMAGE_WRAP_TARGETS.has(value);
}

type ViewSelection = PeerlessViewSnapshot['workerSelection'];
type Position = NonNullable<ViewSelection>['anchor'];

const samePosition = (a: Position, b: Position) => a.story === b.story && a.paraId === b.paraId && a.offset === b.offset;

function sameRange(a: ViewSelection, b: ViewSelection): boolean {
  return !a || !b ? a === b : samePosition(a.anchor, b.anchor) && samePosition(a.focus, b.focus);
}

const NON_EDITING_ACTIONS: ReadonlySet<TextContextAction> = new Set(['selectAll', 'selectTable', 'copy']);

/** Whether a click still acts on what the menu opened for. Sticky position bytes never count; an edit also needs the opened version. */
export function sameContextTarget(action: TextContextAction, opened: PeerlessViewSnapshot, current: PeerlessViewSnapshot): boolean {
  if (opened.identity.session !== current.identity.session || opened.identity.incarnation !== current.identity.incarnation ||
    opened.identity.epoch !== current.identity.epoch) return false;
  if (!NON_EDITING_ACTIONS.has(action) && opened.version !== current.version) return false;
  if (action === 'selectAll') return true;
  if (action === 'selectTable' || action === 'deleteTable') {
    const before = opened.editState?.tableContext;
    const after = current.editState?.tableContext;
    return !!before && !!after && before.story === after.story && before.tableIndex === after.tableIndex;
  }
  return sameRange(opened.workerSelection, current.workerSelection);
}

function isBodyStory(story: string | undefined): boolean {
  return story === undefined || story === 'body' || story.startsWith('body:');
}

export function useWorkerContextMenu(options: {
  binding(): PeerlessEditSessionBinding | null;
  environment(): DocxCommandEnvironment;
  commands: DocxCommandStore;
  clipboard(cut: boolean): Promise<void>;
  submit?(intent: PeerlessPropertyIntent): Promise<DocxCommandResult>;
  onWrapImage?: (handle: string, wrap: ImageLayoutTarget) => void | Promise<boolean>;
  onPaste?: () => void;
  onPasteRefused?: () => void;
  onError?: (error: Error) => void;
  beginAddComment?(): Promise<void>;
  focusEditor?(): void;
}) {
  const [menu, setMenu] = useState<{
    binding: PeerlessEditSessionBinding; snapshot: ReturnType<PeerlessEditSessionBinding['viewStore']['getSnapshot']>;
    position: { x: number; y: number };
    object?: { kind: 'image' | 'cell'; handle: string; version: string; story?: string; x: number; y: number; width: number; height: number };
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
        if (found && rect) { object = { kind, handle: found.handle, version: objects.version, story: found.story, x: rect.x, y: rect.y, width: rect.width, height: rect.height }; break; }
      }
    }
    setMenu({ binding, snapshot, object, position: { x: event.clientX, y: event.clientY } });
  };
  const selection = menu?.snapshot.workerSelection;
  const hasSelection = !!selection && (selection.anchor.story !== selection.focus.story ||
    selection.anchor.paraId !== selection.focus.paraId || selection.anchor.offset !== selection.focus.offset);
  const readOnly = options.environment().readOnly || options.environment().mode === 'viewing';
  const t = options.environment().translate;
  const keys = (id: DocxCommandId) => ({
    shortcut: storeShortcutLabel(options.commands, id, t as never) ?? undefined,
    ariaShortcut: storeAriaShortcut(options.commands, id) ?? undefined,
  });
  const evaluateWrap = (wrap: ImageLayoutTarget, target = menu?.object) => {
    const env = options.environment();
    const currentSelection = options.binding()?.viewStore.getSnapshot().workerSelection ?? selection;
    const bodyStory =
      env.bodyStory !== false &&
      isBodyStory(target?.story) &&
      isBodyStory(selection?.focus.story) &&
      isBodyStory(currentSelection?.focus.story);
    const image = env.image ?? (target?.kind === 'image' ? { wrap: null } : null);
    return evaluateDocxCommand('imageWrap', { wrap }, { ...env, bodyStory, image });
  };
  const items: TextContextMenuItem[] = [
    { action: 'copy', label: t('contextMenu.copy'), disabled: !hasSelection, ...keys('copy') },
    ...readOnly ? [] : [
      { action: 'cut' as const, label: t('contextMenu.cut'), disabled: !hasSelection, ...keys('cut') },
      { action: 'paste' as const, label: t('contextMenu.paste'), ...keys('paste') },
      { action: 'pasteAsPlainText' as const, label: t('contextMenu.pastePlainText'), ...keys('pasteAsPlainText') },
      { action: 'delete' as const, label: t('contextMenu.delete'), disabled: !hasSelection },
    ],
    { action: 'selectAll', label: t('contextMenu.selectAll'), ...keys('selectAll') },
  ];
  if (hasSelection && !readOnly && options.beginAddComment && (selection?.focus.story === 'body' || selection?.focus.story.startsWith('body:'))) {
    items.push({ action: 'addComment', label: t('contextMenu.addComment') });
  }
  if (menu?.object && !readOnly && options.submit) {
    if (menu.object.kind === 'image') {
      items.unshift(
        ...IMAGE_LAYOUT_OPTIONS.map(option => ({
          action: `wrap:${option.choice}` as const,
          label: t(`imageWrap.menu.${option.i18nLabelKey}` as TranslationKey),
          disabled: !evaluateWrap(option.choice).enabled,
        })),
        { action: 'resizeImage', label: t('contextMenu.resizeImage') },
        { action: 'moveImage', label: t('contextMenu.moveImage'), dividerAfter: true },
      );
    } else items.unshift({ action: 'resizeColumn', label: t('contextMenu.columnWidth'), dividerAfter: true });
  }
  if (menu?.snapshot.editState?.tableContext && !readOnly) {
    for (const [action, key] of [['addRowAbove', 'insertRowAbove'], ['addRowBelow', 'insertRowBelow'], ['deleteRow', 'deleteRow'],
      ['addColumnLeft', 'insertColumnLeft'], ['addColumnRight', 'insertColumnRight'], ['deleteColumn', 'deleteColumn'],
      ['mergeCells', 'mergeCells'], ['splitCell', 'splitCell'], ['selectTable', 'selectTable'], ['deleteTable', 'deleteTable']] as const) {
      items.push({ action, label: t(`table.${key}`), disabled: !options.commands.getState('tableAction', action).enabled });
    }
  }
  const pasteClipboard = async (
    current: PeerlessEditSessionBinding,
    action: 'paste' | 'pasteAsPlainText',
    opened: PeerlessViewSnapshot,
    target: PeerlessViewSnapshot['workerSelection'] | undefined,
    clipboardText?: Promise<string>,
  ) => {
    if (action === 'paste') {
      const fragment = await readClipboardFragment();
      if (fragment) {
        if (current !== options.binding()) throw new Error('Document was replaced');
        const latest = current.viewStore.getSnapshot();
        if (!sameContextTarget(action, opened, latest)) throw new Error('Context menu target changed');
        if (options.environment().readOnly || options.environment().mode === 'viewing') return;
        if (isPasteFragmentAllowed(options.commands, fragment)) {
          submitPasteFragment(current.admission, fragment, { source: 'paste', fence: true }, () => options.onPasteRefused?.());
          options.onPaste?.();
          return;
        }
      }
    }
    const text = await (clipboardText ?? navigator.clipboard.readText());
    if (current !== options.binding()) throw new Error('Document was replaced');
    const latest = current.viewStore.getSnapshot();
    if (!sameContextTarget(action, opened, latest)) throw new Error('Context menu target changed');
    if (options.environment().readOnly || options.environment().mode === 'viewing') return;
    const intents = normalizeInputEvent({ source: 'paste', text, eventId: 1, inputTime: performance.now() },
      { composing: false, compositionId: null, lastHandledEventId: null, trailingCompositionText: null }).intents;
    if (target && intents[0]?.kind === 'replaceSelection') intents[0].payload.selection = target;
    admitInputIntents(current.admission, intents, { source: 'paste', fence: true });
    options.onPaste?.();
  };
  const selectAll = (current: PeerlessEditSessionBinding) => {
    current.commands.submit({ kind: 'navigate', payload: { unit: 'document-boundary', direction: 'backward', extend: false } }, { source: 'selection', fence: true });
    current.commands.submit({ kind: 'navigate', payload: { unit: 'document-boundary', direction: 'forward', extend: true } }, { source: 'selection', fence: true });
  };
  /** Runs a clipboard or select-all command from a remapped shortcut. */
  const input = async (id: 'cut' | 'copy' | 'paste' | 'pasteAsPlainText' | 'selectAll', clipboardText?: Promise<string>): Promise<DocxCommandResult> => {
    const current = options.binding();
    if (!current) return { ok: false, failure: { code: 'editor-unavailable', message: options.environment().translate('commands.reasons.commandFailed') } };
    try {
      if (id === 'copy' || id === 'cut') await options.clipboard(id === 'cut');
      else if (id === 'selectAll') selectAll(current);
      else await pasteClipboard(current, id, structuredClone(current.viewStore.getSnapshot()), undefined, clipboardText);
      return { ok: true, status: 'executed' };
    } catch {
      return { ok: false, failure: { code: 'command-failed', message: options.environment().translate('commands.reasons.commandFailed') } };
    }
  };
  const action = async (action: TextContextAction) => {
    if (!menu) return;
    if (action.startsWith('wrap:') && menu.object?.kind === 'image') {
      const handle = menu.object.handle;
      const wrap = action.slice(5);
      close();
      if (!isImageWrapTarget(wrap)) return;
      try {
        const current = options.binding();
        const latest = current?.viewStore.getSnapshot();
        if (current !== menu.binding || latest?.version !== menu.snapshot.version ||
          JSON.stringify(latest.identity) !== JSON.stringify(menu.snapshot.identity)) throw new Error('Context menu target changed');
        const state = evaluateWrap(wrap, menu.object);
        if (!state.enabled) throw new Error(state.disabledReason.message);
        await options.onWrapImage?.(handle, wrap);
      } catch (error) { options.onError?.(error instanceof Error ? error : new Error(String(error))); }
      return;
    }
    if (action === 'resizeImage' || action === 'moveImage' || action === 'resizeColumn') {
      setDialog({ menu, action }); close(); return;
    }
    close();
    try {
      const current = options.binding();
      if (current !== menu.binding) throw new Error('Document was replaced');
      const snapshot = current.viewStore.getSnapshot();
      if (!sameContextTarget(action, menu.snapshot, snapshot)) throw new Error('Context menu target changed');
      if (action === 'copy' || action === 'cut') { await options.clipboard(action === 'cut'); return; }
      if (action === 'selectAll') { selectAll(current); return; }
      if (options.environment().readOnly || options.environment().mode === 'viewing') return;
      if (action === 'addComment') await options.beginAddComment?.();
      else if (action === 'delete' && selection) current.commands.submit({ kind: 'replaceSelection', payload: { text: '', selection } }, { fence: true, source: 'paste' });
      else if (action === 'paste' || action === 'pasteAsPlainText') {
        await pasteClipboard(current, action, menu.snapshot, selection);
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
  return { open, input, element: <>{menu && <TextContextMenu isOpen position={menu.position} hasSelection={hasSelection}
    isEditable={!readOnly} items={items} onAction={(value) => { void action(value); }} onClose={close}
    returnFocus={options.focusEditor} />}
    {dialog?.menu.object && <WorkerObjectDialog action={dialog.action} target={dialog.menu.object} apply={applyObject} close={() => setDialog(null)} />}</> };
}
