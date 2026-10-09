export type LayoutUpdateOrigin = 'local' | 'remote';

export function mergeLayoutUpdateOrigin(
  current: LayoutUpdateOrigin | null,
  next: LayoutUpdateOrigin
): LayoutUpdateOrigin {
  return current === 'local' || next === 'local' ? 'local' : 'remote';
}

export interface ViewportAnchorSnapshot {
  viewportOffset: number;
  scrollTopSnapshot: number;
}

export interface ScrollRestoreTicket<T> {
  generation: number;
  value: T;
}

export class PendingScrollRestoreController<T> {
  private generation = 0;
  private pending: ScrollRestoreTicket<T> | null = null;

  capture(value: T): ScrollRestoreTicket<T> {
    const ticket = { generation: ++this.generation, value };
    this.pending = ticket;
    return ticket;
  }

  cancel(): void {
    this.generation += 1;
    this.pending = null;
  }

  peek(): ScrollRestoreTicket<T> | null {
    return this.pending;
  }

  take(): ScrollRestoreTicket<T> | null {
    const ticket = this.pending;
    this.pending = null;
    return ticket && this.isCurrent(ticket) ? ticket : null;
  }

  isCurrent(ticket: ScrollRestoreTicket<T>): boolean {
    return ticket.generation === this.generation;
  }

  run(ticket: ScrollRestoreTicket<T>, restore: () => void): boolean {
    if (!this.isCurrent(ticket)) return false;
    restore();
    return true;
  }
}

export function computeViewportAnchoredScrollTop(
  anchor: ViewportAnchorSnapshot,
  nextTargetTop: number | null,
  maxScrollTop: number
): number {
  const requested =
    nextTargetTop == null ? anchor.scrollTopSnapshot : nextTargetTop - anchor.viewportOffset;
  return Math.min(Math.max(0, requested), Math.max(0, maxScrollTop));
}

/**
 * A read-only editor's layout passes never move its caret, so only a
 * selection change reveals it there: scrolling back to it on every frame
 * would undo the host's and the reader's navigation.
 */
export function shouldScrollCaretIntoView(
  layoutUpdateOrigin: LayoutUpdateOrigin,
  selectionChanged: boolean,
  readOnly = false
): boolean {
  return (layoutUpdateOrigin === 'local' && !readOnly) || selectionChanged;
}
