import type {
  CommandTicket, EditSession, EditSessionState, LegacyEditSessionOptions,
  PresentationStatus, TicketOutcome,
} from './types';

const MAX_PRESENTATION_WAITERS = 256;
const MAX_DROPPED_SEQS = 256;

type Waiter = { seq: number; resolve(status: PresentationStatus): void };

/** @internal Adapts synchronous legacy edits to render-confirmed command tickets. */
export function createLegacyEditSession<P, S, E, X>(
  options: LegacyEditSessionOptions<P, S, E, X>
): EditSession<P, S, E, X> {
  let state = { ...options.initial };
  let nextSeq = state.throughSeq;
  let nextId = 0;
  let disposed = false;
  let disposal: Promise<void> | undefined;
  const listeners = new Set<(state: EditSessionState<S, E>) => void>();
  const dropped = new Set<number>();
  let waiters: Waiter[] = [];

  function publish(): void {
    const snapshot = state;
    for (const listener of [...listeners]) {
      if (disposed || state !== snapshot) break;
      if (!listeners.has(listener)) continue;
      try { listener(snapshot); } catch {}
    }
  }

  function settle(predicate: (seq: number) => boolean, status: PresentationStatus): void {
    const remaining: Waiter[] = [];
    for (const waiter of waiters) {
      if (predicate(waiter.seq)) waiter.resolve(status);
      else remaining.push(waiter);
    }
    waiters = remaining;
  }

  function drop(seq: number): void {
    dropped.add(seq);
    if (dropped.size > MAX_DROPPED_SEQS) {
      const oldest = dropped.values().next().value;
      if (oldest !== undefined) dropped.delete(oldest);
    }
    settle((waitingSeq) => waitingSeq === seq, 'dropped');
  }

  function whenPresented(seq: number): Promise<PresentationStatus> {
    if (!Number.isSafeInteger(seq) || seq < 0 || dropped.has(seq)) return Promise.resolve('dropped');
    if (seq <= state.presentedThroughSeq) return Promise.resolve('presented');
    if (disposed) return Promise.resolve('dropped');
    if (waiters.length >= MAX_PRESENTATION_WAITERS) drop(waiters[0]!.seq);
    if (dropped.has(seq)) return Promise.resolve('dropped');
    return new Promise((resolve) => { waiters.push({ seq, resolve }); });
  }

  function confirm(coversThroughSeq: number): void {
    if (disposed || !Number.isSafeInteger(coversThroughSeq) ||
      coversThroughSeq <= state.presentedThroughSeq || coversThroughSeq > state.throughSeq) return;
    state = { ...state, presentedThroughSeq: coversThroughSeq };
    settle((seq) => seq <= coversThroughSeq, 'presented');
    publish();
  }

  function ticket(
    id: number, seq: number | undefined, outcome: TicketOutcome,
    presented: Promise<PresentationStatus>
  ): CommandTicket {
    const ack = Promise.resolve(outcome);
    void ack.catch(() => {});
    return { id, get seq() { return seq; }, ack, presented };
  }

  function dispose(): Promise<void> {
    if (disposal) return disposal;
    disposed = true;
    listeners.clear();
    settle(() => true, 'dropped');
    dropped.clear();
    let resolve!: () => void;
    let reject!: (error: unknown) => void;
    disposal = new Promise((done, fail) => { resolve = done; reject = fail; });
    void disposal.catch(() => {});
    try {
      Promise.resolve(options.dispose?.()).then(resolve, reject);
    } catch (error) {
      reject(error);
    }
    return disposal;
  }

  return {
    backend: options.backend,
    get state() { return state; },
    submit(kind, payload, submitOptions) {
      const id = ++nextId;
      if (disposed) {
        return ticket(id, undefined, { status: 'cancelled', reason: 'disposed' },
          Promise.resolve('dropped'));
      }
      const seq = ++nextSeq;
      const result = options.apply(kind, payload, submitOptions);
      state = { ...result.state, throughSeq: Math.max(seq, result.state.throughSeq),
        presentedThroughSeq: state.presentedThroughSeq };
      nextSeq = Math.max(nextSeq, state.throughSeq);
      const refused = result.outcome.status === 'refused' || result.outcome.status === 'cancelled';
      if (refused) drop(seq);
      const submitted = ticket(id, seq, { ...result.outcome, seq },
        refused ? Promise.resolve('dropped') : whenPresented(seq));
      publish();
      if (!disposed) options.render(seq, confirm);
      return submitted;
    },
    whenPresented,
    subscribe(listener) {
      if (!disposed) listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    flush: () => Promise.resolve(),
    dispose,
  };
}
