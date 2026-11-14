import { useCallback, useEffect, type RefObject } from 'react';
import { indexGeometry } from '@betteroffice/pptx/worker';
import { useContextMenu } from '../../../../shared/react-toolbar/ContextMenu';
import type { OverflowMenuEntry } from '../../../../shared/react-toolbar/MenuList';
import type { PptxCommandStore } from '../commands/types';
import { pptxCommandController } from '../commands/createPptxCommandStore';
import { pptxContextMenuEntries, type PptxContextMenuTarget } from '../contextMenu/menus';
import { useMenuTranslation } from '../i18n/useMenuTranslation';
import { hitConfirmedGeometry, peerlessKnownState } from './peerlessViewStore';
import type { PeerlessPresentation } from './usePeerlessPresentation';

function disable(entries: readonly OverflowMenuEntry[]): OverflowMenuEntry[] {
  return entries.map((entry) => entry.kind === 'item' ? { ...entry, disabled: true }
    : entry.kind === 'submenu' ? { ...entry, disabled: true, entries: disable(entry.entries) }
      : entry.kind === 'group' ? { ...entry, entries: disable(entry.entries) } : entry);
}

/** @internal */
export function usePeerlessContextMenu(run: PeerlessPresentation | null, commands: PptxCommandStore,
  surface: RefObject<HTMLDivElement | null>, canvas: RefObject<HTMLCanvasElement | null>, reviewing: boolean,
  focus: () => void) {
  const { t } = useMenuTranslation();
  const subscribe = useCallback((listener: () => void) => {
    const releases = [commands.subscribe(listener), run?.store.subscribePresentation(listener)];
    return () => { for (const release of releases) release?.(); };
  }, [commands, run]);
  const menu = useContextMenu({
    surface,
    subscribe,
    keyboardPoint() {
      const view = run?.store.getSnapshot();
      const bounds = canvas.current?.getBoundingClientRect();
      const frame = view?.confirmed;
      if (!view || !bounds || !frame) return null;
      const rect = view.selection?.kind === 'object' ? view.selection.bounds
        : view.selectionRects[0] ?? view.caret;
      if (!rect) return null;
      return { x: bounds.left + rect.x * bounds.width / frame.displayList.width,
        y: bounds.top + (rect.y + rect.height) * bounds.height / frame.displayList.height };
    },
    resolve(request) {
      if (!run?.current || !run.selectionEnabled || reviewing ||
        request.target instanceof Element && request.target.closest('[role="menu"]')) return null;
      const view = run.store.getSnapshot();
      pptxCommandController(commands)?.refresh();
      const frame = view.confirmed;
      const bounds = canvas.current?.getBoundingClientRect();
      if (!frame || !bounds?.width || !bounds.height || view.composition) return null;
      let target: PptxContextMenuTarget;
      let ready = () => true;
      if (request.modality === 'pointer') {
        const point = { x: (request.x - bounds.left) * frame.displayList.width / bounds.width,
          y: (request.y - bounds.top) * frame.displayList.height / bounds.height };
        if (point.x < 0 || point.y < 0 || point.x > frame.displayList.width || point.y > frame.displayList.height) return null;
        const hit = hitConfirmedGeometry(run.store, point, bounds.width / frame.displayList.width);
        if (!hit) return null;
        const selected = view.selection;
        if (hit.kind === 'text') {
          target = indexGeometry(frame.displayList).hits.some(({ primitive, clips }) =>
            primitive.shapeId === hit.shapeId && primitive.kind === 'textBox' && primitive.storyId === hit.storyId &&
            clips.some((clip) => clip.kind === 'table')) ? 'table' : 'text';
          const same = selected && selected.kind !== 'object' && selected.kind !== 'shape' &&
            selected.slideId === frame.frame.slideId && selected.shapeId === hit.shapeId && selected.storyId === hit.storyId &&
            hit.position >= Math.min(selected.anchor, selected.focus) && hit.position <= Math.max(selected.anchor, selected.focus);
          if (!same) {
            ready = () => {
              const next = run.store.getSnapshot().selection;
              return Boolean(next && next.kind !== 'object' && next.slideId === frame.frame.slideId &&
                next.shapeId === hit.shapeId && next.storyId === hit.storyId && next.anchor === hit.position && next.focus === hit.position);
            };
            void run.store.dispatch({ kind: 'selectAtPoint', payload: { frame: frame.frame, ...point, extend: false,
              inputTime: performance.now(), undoGroup: 'selection' } }).catch(() => {});
          }
        } else {
          target = 'shape';
          ready = () => {
            const next = run.store.getSnapshot().selection;
            return next?.kind === 'object' && next.slideId === frame.frame.slideId && next.shapeId === hit.shapeId;
          };
          if (!ready()) void run.store.dispatch({ kind: 'selectObject', payload: { frame: frame.frame, shapeId: hit.shapeId,
            inputTime: performance.now(), undoGroup: 'selection' } }).catch(() => {});
        }
      } else {
        if (!view.selection) return null;
        target = view.selection.kind === 'object' || view.selection.kind === 'shape' ? 'shape'
          : view.editState.toolbar.table ? 'table' : 'text';
      }
      return { label: t(target === 'table' ? 'contextMenu.table' : target === 'shape' ? 'contextMenu.shape' : 'contextMenu.text'),
        returnFocus: focus,
        entries: () => {
          const entries = pptxContextMenuEntries(target, commands, t);
          return ready() && peerlessKnownState(run.store) ? entries : disable(entries);
        } };
    },
  });
  useEffect(() => menu.close(), [run, reviewing, menu.close]);
  return menu;
}
