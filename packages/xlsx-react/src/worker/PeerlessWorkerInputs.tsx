import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { commandForEvent } from '../commands/descriptors';
import { FormulaBarContext, presentInputValue, type FormulaBarBinding } from '../components/toolbar/FormulaBar';
import type { WorkerDraftController } from './WorkerDraftController';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';
import type { WorkerPresentationStore } from './WorkerPresentationStore';

export const PeerlessWorkerInputs = memo(function PeerlessWorkerInputs({ controller, presentation, run, active,
  selectedChart, error, readOnly, navigation, children }: {
  controller: WorkerDraftController;
  presentation: WorkerPresentationStore;
  run: PeerlessWorkbookSession | null;
  active: number;
  selectedChart: boolean;
  error: boolean;
  readOnly: boolean;
  navigation: { current: number };
  children: ReactNode;
}) {
  const { selection, view, painted } = useSyncExternalStore(presentation.subscribe, presentation.getSnapshot);
  const [focusedInput, setFocusedInput] = useState<{ at: string; value: string } | null>(null);
  const inputRevision = painted?.coversThroughSeq ?? 0;
  const unavailable = usePeerlessReadOnly(run, readOnly);
  useEffect(() => {
    if (!run?.current || run.failure || !selection || selectedChart) return;
    const identity = `${run.generation}:${run.session.confirmedBasis.epoch}`;
    const at = navigation.current;
    const draftIdentity = controller.focusedCellIdentity;
    const cell = `${active}:${cellAddress(selection.anchor)}`;
    let cancelled = false;
    void run.session.call.cellInputs(active, cellAddress(selection.anchor)).then((result) => {
      if (!cancelled && run.current && !run.failure && identity === `${run.generation}:${run.session.confirmedBasis.epoch}` &&
        navigation.current === at && controller.focusedCellIdentity === draftIdentity) {
        const value = result.cells[0]?.[0]?.input ?? '';
        controller.focusedCellInput = { identity: draftIdentity, value };
        setFocusedInput({ at: cell, value });
        controller.presented(inputRevision);
      }
    }).catch((failure) => { if (!cancelled && run.current) run.fail(failure); });
    return () => { cancelled = true; };
  }, [run, active, selection, selectedChart, painted, inputRevision, controller, navigation]);
  useLayoutEffect(() => { setFocusedInput(null); }, [run]);
  const state = useSyncExternalStore(controller.subscribe, controller.getSnapshot);
  const handlers = useMemo<Omit<FormulaBarBinding, 'a1' | 'value' | 'disabled' | 'readOnly'>>(() => ({
    native: true,
    inputRef: (input) => { controller.formulaInput = input; },
    isComposing: () => controller.isComposing('formula'),
    onChange: (value) => controller.change('formula', value, controller.formulaInput),
    onCommit: (move) => controller.commit('formula', move), onCancel: controller.cancel,
    onBlur: () => controller.commit('formula'),
    onCompositionStart: () => controller.start('formula'), onCompositionEnd: () => controller.end(),
    onKeyDown: (event) => {
      const consumed = controller.key('formula', event);
      if (consumed) { event.stopPropagation(); return true; }
      return false;
    },
    onKeyUp: () => controller.keyUp(),
  }), [controller]);
  const a1 = selection ? cellAddress(selection.anchor) : '';
  const value = state.formula ?? (focusedInput?.at === `${active}:${a1}` ? focusedInput.value : '');
  useLayoutEffect(() => {
    const input = controller.formulaInput;
    if (input && !controller.isComposing('formula') && input.value !== value) presentInputValue(input, value);
  }, [controller, state, value]);
  const disabled = !view || !selection || error;
  const formula = useMemo<FormulaBarBinding>(() => ({ ...handlers, a1, value, disabled, readOnly: unavailable || selectedChart }),
    [handlers, a1, value, disabled, unavailable, selectedChart]);
  return <FormulaBarContext.Provider value={formula}>{children}
    {state.refused.map((record) => <button key={record.seq ?? record.draft.revision}
      data-testid="xlsx-refused-draft" data-sheet={record.draft.sheet} data-row={record.draft.row}
      data-col={record.draft.col} onClick={() => controller.restore(record)}>
      {record.draft.value}
    </button>)}
  </FormulaBarContext.Provider>;
});

export const PeerlessWorkerCellInput = memo(function PeerlessWorkerCellInput({ controller, rect, disabled, zoom }: {
  controller: WorkerDraftController;
  rect(row: number, col: number, sheet: number): { x: number; y: number; w: number; h: number } | null;
  disabled: boolean;
  zoom: number;
}) {
  const inputRef = useCallback((input: HTMLInputElement | null) => { controller.cellInput = input; }, [controller]);
  const state = useSyncExternalStore(controller.subscribe, controller.getSnapshot);
  const draft = state.cell;
  const retained = useRef<{ target: string; rect: ReturnType<typeof rect> }>({ target: '', rect: null });
  const target = draft ? `${draft.generation}:${draft.sheet}:${draft.row}:${draft.col}` : '';
  if (target !== retained.current.target) retained.current = { target, rect: null };
  const current = draft ? rect(draft.row, draft.col, draft.sheet) : null;
  if (current && !controller.isComposing('cell')) retained.current.rect = current;
  const bounds = retained.current.rect;
  useLayoutEffect(() => {
    const input = controller.cellInput;
    if (!input) return;
    const beforeInput = (event: InputEvent) => {
      if (controller.holdCellBeforeInput(event)) event.preventDefault();
    };
    const select = () => controller.captureCellSelection();
    input.addEventListener('beforeinput', beforeInput);
    input.addEventListener('select', select);
    return () => {
      input.removeEventListener('beforeinput', beforeInput);
      input.removeEventListener('select', select);
    };
  }, [controller, draft !== null, bounds !== null]);
  useLayoutEffect(() => {
    const input = controller.cellInput;
    if (!draft || !input || controller.isComposing('cell')) return;
    if (input.value !== draft.value) input.value = draft.value;
    const active = input.ownerDocument.activeElement;
    if (active !== input) {
      const surface = input.closest('[tabindex]');
      if (active && active !== input.ownerDocument.body && !surface?.contains(active)) return;
      input.focus({ preventScroll: true }); input.setSelectionRange(input.value.length, input.value.length);
    }
    controller.restoreCellSelection();
  }, [controller, draft, bounds !== null]);
  if (!draft) return null;
  if (!bounds) return null;
  const style: CSSProperties = { position: 'absolute', left: bounds.x * zoom, top: bounds.y * zoom,
    width: bounds.w * zoom, height: bounds.h * zoom, boxSizing: 'border-box', border: '2px solid #217346',
    background: '#ffffff', padding: '0 3px', font: `${13 * zoom}px system-ui, sans-serif`,
    pointerEvents: 'auto', outline: 'none' };
  return <input ref={inputRef} data-testid="xlsx-cell-editor"
    defaultValue={draft.value} disabled={disabled} style={style}
    onChange={(event) => controller.change('cell', event.target.value, event.currentTarget)}
    onCompositionStart={() => controller.start('cell')}
    onCompositionEnd={(event) => controller.end(event.nativeEvent.data)}
    onKeyDown={(event) => {
      if (!commandForEvent(event)) event.stopPropagation();
      if (controller.key('cell', event)) return;
      if (controller.holdCellKey(event)) { event.preventDefault(); event.stopPropagation(); return; }
      if (event.key === 'Enter') { controller.commit('cell', event.shiftKey ? 'up' : 'down'); event.preventDefault(); }
      else if (event.key === 'Tab') { controller.commit('cell', event.shiftKey ? 'left' : 'right'); event.preventDefault(); }
      else if (event.key === 'Escape') { controller.cancel(); event.preventDefault(); }
    }}
    onKeyUp={() => controller.keyUp()} onBlur={() => controller.commit('cell')} />;
});

function cellAddress(cell: { row: number; col: number }): string {
  let column = '';
  for (let col = cell.col + 1; col > 0; col = Math.floor((col - 1) / 26)) {
    column = String.fromCharCode(65 + (col - 1) % 26) + column;
  }
  return `${column}${cell.row + 1}`;
}

export function usePeerlessReadOnly(run: PeerlessWorkbookSession | null, readOnly: boolean): boolean {
  const subscribe = useCallback((listener: () => void) => run?.subscribe(listener) ?? (() => {}), [run]);
  const getSnapshot = useCallback(() => !run || !run.ready || run.readOnly, [run]);
  return useSyncExternalStore(subscribe, getSnapshot) || readOnly;
}
