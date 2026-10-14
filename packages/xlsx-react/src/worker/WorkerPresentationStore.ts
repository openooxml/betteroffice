import { useMemo, useSyncExternalStore } from 'react';
import type { Selection, WorkbookSheetView } from '@betteroffice/xlsx';
import type { WorkerPaintResult } from './PeerlessWorkerPaintSource';

interface PresentationSnapshot {
  selection: Selection | null;
  view: WorkbookSheetView | null;
  painted: WorkerPaintResult | null;
}

function sameGeometry(previous: WorkbookSheetView | null, next: WorkbookSheetView | null): boolean {
  return previous === next || Boolean(previous && next && previous.sheet === next.sheet &&
    previous.contentWidth === next.contentWidth && previous.contentHeight === next.contentHeight &&
    previous.frozenRows === next.frozenRows && previous.frozenCols === next.frozenCols &&
    previous.frozenWidth === next.frozenWidth && previous.frozenHeight === next.frozenHeight &&
    previous.initialScrollX === next.initialScrollX && previous.initialScrollY === next.initialScrollY);
}

export class WorkerPresentationStore {
  readonly selectionRef: { current: Selection | null } = { current: null };
  readonly viewRef: { current: WorkbookSheetView | null } = { current: null };
  private snapshot: PresentationSnapshot = { selection: null, view: null, painted: null };
  private listeners = new Set<() => void>();
  private publication: object | null = null;

  getSnapshot = (): PresentationSnapshot => this.snapshot;
  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  };
  private publish = (): void => {
    this.publication = null;
    for (const listener of this.listeners) listener();
  };
  private schedule(): void {
    if (this.publication) return;
    const publication = {};
    this.publication = publication;
    queueMicrotask(() => { if (this.publication === publication) this.publish(); });
  }
  setSelection = (selection: Selection | null): void => {
    this.selectionRef.current = selection;
    if (this.snapshot.selection === selection) return;
    this.snapshot = { ...this.snapshot, selection };
    this.publish();
  };
  setView = (view: WorkbookSheetView | null): void => {
    this.viewRef.current = view;
    if (sameGeometry(this.snapshot.view, view)) return;
    this.snapshot = { ...this.snapshot, view };
    this.schedule();
  };
  setPainted = (painted: WorkerPaintResult | null): void => {
    if (this.snapshot.painted === painted) return;
    this.snapshot = { ...this.snapshot, painted };
    this.schedule();
  };
  reset(): void {
    this.selectionRef.current = null;
    this.viewRef.current = null;
    this.snapshot = { selection: null, view: null, painted: null };
    this.publish();
  }
  cancel(): void {
    this.publication = null;
  }
}

export function useWorkerPresentation<T>(store: WorkerPresentationStore, select: (state: PresentationSnapshot) => T): T {
  const getSnapshot = useMemo(() => {
    let previous: { state: PresentationSnapshot; value: T } | undefined;
    return () => {
      const state = store.getSnapshot();
      if (!previous || previous.state !== state) previous = { state, value: select(state) };
      return previous.value;
    };
  }, [store, select]);
  return useSyncExternalStore(store.subscribe, getSnapshot);
}
