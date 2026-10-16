import type {
  CommandTicket, EchoEntry, EchoHooks, FrameDisposition, FrameEnvelope,
  PresentationStatus, Presenter, PresenterOptions,
} from './types';

type Waiter = {
  promise: Promise<PresentationStatus>;
  resolve(status: PresentationStatus): void;
};

const MAX_WAITERS = 1024;
const MAX_DROPPED_SEQS = 256;

/** @internal Creates a bounded echo reconciler with synchronous frame installation and coverage waiters. */
export function createPresenter<X, F>(options: PresenterOptions<X, F>): Presenter<X, F> {
  const echoLimit = Number.isSafeInteger(options.maxEchoEntries) && options.maxEchoEntries! >= 0
    ? options.maxEchoEntries! : 64;
  const now = options.now ?? Date.now;
  const waiters = new Map<number, Waiter>();
  const dropped = new Set<number>();
  let pending: EchoEntry<X>[] = [];
  let coverage = -1;
  let settledCoverage = -1;
  let paintedCoverage = -1;
  let revision: number | undefined;
  let installing: { coverage: number; revision: number } | undefined;
  let resetGeneration = 0;
  let disposed = false;
  let eventDepth = 0;
  let eventTotalMs = 0;

  function event<T>(callback: () => T): T {
    if (eventDepth === 0) eventTotalMs = 0;
    eventDepth += 1;
    try { return callback(); } finally { eventDepth -= 1; }
  }

  function paint(hook: keyof EchoHooks<X, F>, callback: () => void): boolean {
    const started = time();
    const previousTotalMs = eventTotalMs;
    try { callback(); return true; } catch { return false; } finally {
      const elapsedMs = Math.max(0, time() - started);
      eventTotalMs = previousTotalMs + elapsedMs;
      if (eventTotalMs > 2) {
        try { options.onEchoOverrun?.({ hook, elapsedMs, eventTotalMs }); } catch {}
      }
    }
  }

  function time(): number {
    try { const value = now(); return Number.isFinite(value) ? value : 0; }
    catch { return 0; }
  }

  function settle(seq: number, status: PresentationStatus): void {
    const waiter = waiters.get(seq);
    if (!waiter) return;
    waiters.delete(seq);
    waiter.resolve(status);
  }

  function recordDrop(seq: number): void {
    if (!dropped.has(seq)) {
      if (dropped.size >= MAX_DROPPED_SEQS) dropped.delete(dropped.values().next().value!);
      dropped.add(seq);
    }
    settle(seq, 'dropped');
  }

  function add(entry: EchoEntry<X>): boolean {
    return event(() => {
      if (disposed || echoLimit === 0) return false;
      if (pending.length >= echoLimit) {
        if (options.echoOverflow !== 'evict-oldest') return false;
        pending = pending.slice(pending.length - echoLimit + 1);
      }
      pending = [...pending, entry];
      if (options.echo) paint('paint', () => options.echo!.paint(pending));
      return true;
    });
  }

  function replace(tickets: readonly CommandTicket[], entry?: EchoEntry<X>): void {
    event(() => {
      if (disposed) return;
      const ids = new Set(tickets.map((ticket) => ticket.id));
      let remaining = pending.filter((echo) => !ids.has(echo.ticket.id));
      const changed = remaining.length !== pending.length;
      if (entry !== undefined && echoLimit > 0 && remaining.length >= echoLimit &&
        options.echoOverflow === 'evict-oldest') {
        remaining = remaining.slice(remaining.length - echoLimit + 1);
      }
      const accepted = entry !== undefined && remaining.length < echoLimit;
      pending = accepted ? [...remaining, entry!] : remaining;
      if ((changed || accepted) && options.echo) paint('paint', () => options.echo!.paint(pending));
    });
  }

  function drop(tickets: readonly CommandTicket[], reason: string): void {
    event(() => {
      if (disposed) return;
      const ids = new Set(tickets.map((ticket) => ticket.id));
      const removed = pending.filter((entry) => ids.has(entry.ticket.id));
      pending = pending.filter((entry) => !ids.has(entry.ticket.id));
      for (const entry of removed) {
        const seq = entry.seq ?? entry.ticket.seq;
        if (seq !== undefined) recordDrop(seq);
      }
      for (const ticket of tickets) if (ticket.seq !== undefined) recordDrop(ticket.seq);
      if (removed.length > 0 && options.echo) {
        paint('drop', () => options.echo!.drop(removed, reason));
      }
    });
  }

  function presentTransactional(
    frame: FrameEnvelope<F>, sameCoverage: boolean,
    install: (retired: readonly EchoEntry<X>[], remaining: readonly EchoEntry<X>[]) => boolean
  ): FrameDisposition {
    const retired: EchoEntry<X>[] = [];
    const remaining: EchoEntry<X>[] = [];
    for (const entry of pending) {
      const seq = entry.seq ?? entry.ticket.seq;
      (!sameCoverage && seq !== undefined && seq <= frame.coversThroughSeq ? retired : remaining).push(entry);
    }
    const previousInstalling = installing;
    const generation = resetGeneration;
    installing = { coverage: frame.coversThroughSeq, revision: frame.revision };
    try {
      if (!install(retired, remaining) || disposed || generation !== resetGeneration) return 'stale';
      coverage = Math.max(coverage, frame.coversThroughSeq);
      revision = revision === undefined ? frame.revision : Math.max(revision, frame.revision);
      if (sameCoverage) return 'presented';
      const retiring = new Set(retired);
      pending = pending.filter((entry) => !retiring.has(entry));
      const throughSeq = Math.min(coverage, frame.coversThroughSeq);
      paintedCoverage = Math.max(paintedCoverage, throughSeq);
      for (const seq of waiters.keys()) {
        if (seq <= throughSeq) settle(seq, dropped.has(seq) ? 'dropped' : 'presented');
      }
      settledCoverage = Math.max(settledCoverage, throughSeq);
      return 'presented';
    } finally {
      if (generation === resetGeneration) installing = previousInstalling;
    }
  }

  function present(frame: FrameEnvelope<F>): FrameDisposition {
    return event(() => {
      let retained = false;
      function install(retired: readonly EchoEntry<X>[], remaining: readonly EchoEntry<X>[]): boolean {
        if (!options.echo) { retained = true; return true; }
        const painted = paint('present', () => {
          const result = options.echo!.present(frame, retired, remaining);
          retained = !options.transactionalPresentation || result !== false;
        });
        return painted && retained;
      }
      try {
        const sameCoverage = frame.coversThroughSeq === coverage;
        if (disposed || frame.coversThroughSeq < coverage || (sameCoverage &&
          (!options.sameCoverageFrames || coverage < 0 || revision === undefined || frame.revision <= revision))) return 'stale';
        if (installing && (frame.coversThroughSeq < installing.coverage ||
          (frame.coversThroughSeq === installing.coverage &&
            (!options.sameCoverageFrames || frame.revision <= installing.revision)))) return 'stale';
        if (frame.baseRevision !== undefined && frame.baseRevision !== revision) {
          options.requestFullFrame();
          return 'base-mismatch';
        }
        if (options.transactionalPresentation && options.echo) {
          return presentTransactional(frame, sameCoverage, install);
        }
        if (sameCoverage) {
          revision = frame.revision;
          install([], pending);
          if (disposed) return 'stale';
          return 'presented';
        }
        const retired: EchoEntry<X>[] = [];
        const remaining: EchoEntry<X>[] = [];
        const previousCoverage = coverage;
        const previousPaintedCoverage = paintedCoverage;
        for (const entry of pending) {
          const seq = entry.seq ?? entry.ticket.seq;
          (seq !== undefined && seq <= frame.coversThroughSeq ? retired : remaining).push(entry);
        }
        pending = remaining;
        coverage = frame.coversThroughSeq;
        revision = frame.revision;
        const painted = install(retired, remaining);
        if (disposed) return 'stale';
        const throughSeq = Math.min(coverage, frame.coversThroughSeq);
        if (painted) paintedCoverage = Math.max(paintedCoverage, throughSeq);
        if (!painted) {
          for (const entry of retired) {
            const seq = entry.seq ?? entry.ticket.seq;
            if (seq !== undefined && seq <= frame.coversThroughSeq &&
              (paintedCoverage === previousPaintedCoverage || seq > paintedCoverage)) recordDrop(seq);
          }
        }
        for (const seq of waiters.keys()) {
          if (seq > throughSeq) continue;
          if (!painted) {
            if (seq > previousCoverage) recordDrop(seq);
          } else settle(seq, dropped.has(seq) ? 'dropped' : 'presented');
        }
        if (!painted) {
          for (let seq = Math.max(previousCoverage + 1, paintedCoverage + 1, throughSeq - MAX_DROPPED_SEQS + 1);
            seq <= throughSeq; seq += 1) {
            recordDrop(seq);
          }
        }
        settledCoverage = Math.max(settledCoverage, throughSeq);
        return 'presented';
      } finally {
        if (!retained) {
          try { options.discardFrame?.(frame); } catch {}
        }
      }
    });
  }

  function whenPresented(seq: number): Promise<PresentationStatus> {
    if (disposed || dropped.has(seq)) return Promise.resolve('dropped');
    if (seq <= settledCoverage) return Promise.resolve('presented');
    const existing = waiters.get(seq);
    if (existing) return existing.promise;
    if (waiters.size >= MAX_WAITERS) recordDrop(waiters.keys().next().value!);
    let resolve!: (status: PresentationStatus) => void;
    const promise = new Promise<PresentationStatus>((done) => { resolve = done; });
    waiters.set(seq, { promise, resolve });
    return promise;
  }

  return {
    get presentedThroughSeq() { return coverage; },
    get revision() { return revision; },
    get pending() { return pending.slice(); },
    event,
    add,
    replace,
    drop,
    present,
    whenPresented,
    reset() {
      if (disposed) return;
      resetGeneration += 1;
      installing = undefined;
      coverage = -1;
      settledCoverage = -1;
      paintedCoverage = -1;
      revision = undefined;
    },
    dispose() {
      event(() => {
        if (disposed) return;
        disposed = true;
        const removed = pending;
        pending = [];
        for (const seq of waiters.keys()) settle(seq, 'dropped');
        dropped.clear();
        if (removed.length > 0 && options.echo) {
          paint('drop', () => options.echo!.drop(removed, 'disposed'));
        }
      });
    },
  };
}
