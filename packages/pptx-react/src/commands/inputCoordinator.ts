/** Raised by work queued behind input that failed. */
export class PptxInputFailure extends Error {
  constructor(readonly cause: unknown) {
    super(cause instanceof Error ? cause.message : String(cause));
  }
}

/** Raised instead of running work queued for a document that has since been replaced. */
export class PptxInputStale extends Error {
  constructor() {
    super('The presentation was replaced before queued work ran');
  }
}

export type PptxInputKind = 'keyboard' | 'other';

/**
 * Orders editor input and the commands issued between it: work runs at once
 * while nothing is pending, otherwise after everything accepted before it.
 */
export interface PptxInputCoordinator {
  /** Whether accepted input or commands are still waiting. */
  busy(): boolean;
  /** Whether queued keyboard input will move the selection before the next command. */
  keyboardQueued(): boolean;
  /** Applies input now when idle, otherwise after everything accepted before it. */
  input(work: () => void | Promise<void>, kind?: PptxInputKind): Promise<void>;
  /** Runs `operation` after everything accepted before; rejects with {@link PptxInputFailure} when that input failed. */
  run<T>(operation: () => T | Promise<T>): Promise<T>;
  /** Abandons the queue for a newly opened document; queued work then rejects with {@link PptxInputStale}. */
  reset(): void;
}

function isThenable(value: unknown): value is PromiseLike<unknown> {
  return (
    value !== null &&
    (typeof value === 'object' || typeof value === 'function') &&
    typeof (value as PromiseLike<unknown>).then === 'function'
  );
}

const noop = () => {};

/** Creates a coordinator; `onBusyChange` fires when the queue fills or drains. */
export function createInputCoordinator(onBusyChange: () => void = noop): PptxInputCoordinator {
  let tail: Promise<void> = Promise.resolve();
  let pending = 0;
  let keyboard = 0;
  let epoch = 0;
  let unsettled: Promise<unknown>[] = [];
  const failures = new Map<Promise<unknown>, unknown>();

  const track = (item: Promise<unknown>, kind: PptxInputKind | null) => {
    const itemEpoch = epoch;
    pending += 1;
    if (kind === 'keyboard') keyboard += 1;
    if (kind) unsettled.push(item);
    if (pending === 1) onBusyChange();
    const settle = (error?: { value: unknown }) => {
      if (itemEpoch !== epoch) return;
      if (error && kind) failures.set(item, error.value);
      unsettled = unsettled.filter((candidate) => candidate !== item);
      pending -= 1;
      if (kind === 'keyboard') keyboard -= 1;
      if (pending === 0) {
        failures.clear();
        onBusyChange();
      }
    };
    item.then(
      () => settle(),
      (value: unknown) => settle({ value })
    );
  };

  const enqueue = <T>(work: () => T | Promise<T>): Promise<T> => {
    const itemEpoch = epoch;
    const result = tail.then(() => {
      if (itemEpoch !== epoch) throw new PptxInputStale();
      return work();
    });
    tail = result.then(noop, noop);
    return result;
  };

  return {
    busy: () => pending > 0,
    keyboardQueued: () => keyboard > 0,
    input(work, kind = 'other') {
      if (pending === 0) {
        let value: void | Promise<void>;
        try {
          value = work();
        } catch (error) {
          return Promise.reject(error);
        }
        if (!isThenable(value)) return Promise.resolve();
        const item = Promise.resolve(value);
        tail = item.then(noop, noop);
        track(item, kind);
        return item;
      }
      const item = enqueue(work);
      track(item, kind);
      return item;
    },
    run(operation) {
      const ahead = [...unsettled];
      const checked = () => {
        for (const item of ahead) {
          if (failures.has(item)) throw new PptxInputFailure(failures.get(item));
        }
        return operation();
      };
      if (pending === 0) {
        let value: ReturnType<typeof operation>;
        try {
          value = operation();
        } catch (error) {
          return Promise.reject(error);
        }
        if (!isThenable(value)) return Promise.resolve(value);
        const item = Promise.resolve(value);
        tail = item.then(noop, noop);
        track(item, null);
        return item;
      }
      const item = enqueue(checked);
      track(item, null);
      return item;
    },
    reset() {
      epoch += 1;
      const wasBusy = pending > 0;
      tail = Promise.resolve();
      pending = 0;
      keyboard = 0;
      unsettled = [];
      failures.clear();
      if (wasBusy) onBusyChange();
    },
  };
}
