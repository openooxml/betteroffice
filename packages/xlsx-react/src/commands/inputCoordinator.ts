import { XlsxCommandAdmissionError } from './createXlsxCommandStore';

/** Cell text typed into the in-cell editor or the formula bar and not yet written. */
export interface InputDraft {
  readonly generation: number;
  readonly sheet: number;
  readonly row: number;
  readonly col: number;
  readonly value: string;
  readonly source: 'cell' | 'formula';
}

/** What a command waits for before it is admitted. */
export interface InputSeal {
  refused?: 'gesture-active';
  /** True once an IME composition in progress ends, false if its input goes away first. */
  composition?: Promise<boolean>;
}

export interface InputCoordinatorHooks {
  generation(): number;
  /** Lands synchronous input such as an arrow-key burst and ends composition. */
  seal(): InputSeal;
  /** Reads the text the draft's input shows into the draft. */
  sync(): void;
  /** Writes `draft`; false leaves it unwritten. */
  write(draft: InputDraft): boolean;
  /** Closes the input of a written draft that is still the live one. */
  close(draft: InputDraft): void;
  /** Shows a rejected draft in its own input when its sheet is active; false otherwise. */
  restore(draft: InputDraft): boolean;
}

/**
 * Orders every input write and the commands admitted after it in one queue.
 * A draft whose write failed stays rejected, at its own cell, until it is
 * written or discarded; until then every command fails with `input-failed`.
 */
export interface InputCoordinator {
  readonly draft: InputDraft | null;
  readonly rejected: readonly InputDraft[];
  /** Whether accepted input is queued and not yet written. */
  readonly pending: boolean;
  setDraft(draft: InputDraft | null): void;
  /** Queues an input write; a failure (throw or false) fails the commands behind it. */
  input(write: () => boolean | void | Promise<boolean | void>): Promise<void>;
  /** Commits a finished draft in order; false only when written at once and failed. */
  submit(draft: InputDraft): boolean;
  /** Drops the rejected draft of the same cell and input. */
  discard(draft: InputDraft): void;
  /**
   * Closes the live draft for a synchronous host call: it is written now when
   * nothing is queued, else queued behind that input. False when it was written
   * at once and refused for the first time.
   */
  settle(): boolean;
  /** Runs `operation` once the input accepted before this call is written. */
  runAfterPendingInput<T>(operation: () => T | Promise<T>): Promise<T>;
}

function isPromise(value: unknown): value is Promise<unknown> {
  return typeof (value as Promise<unknown> | null)?.then === 'function';
}

function sameCell(a: InputDraft, b: InputDraft): boolean {
  return (
    a.generation === b.generation &&
    a.source === b.source &&
    a.sheet === b.sheet &&
    a.row === b.row &&
    a.col === b.col
  );
}

export function createInputCoordinator(hooks: InputCoordinatorHooks): InputCoordinator {
  let draft: InputDraft | null = null;
  let tail: Promise<void> | null = null;
  let failed = false;
  let sealing = false;
  let sealFailed = false;
  const written = new WeakSet<InputDraft>();
  let rejected: readonly InputDraft[] = [];

  const settleRejected = (target: InputDraft) =>
    (rejected = rejected.filter((entry) => !sameCell(entry, target)));
  const reject = (target: InputDraft) =>
    (rejected = [...rejected.filter((entry) => !sameCell(entry, target)), target]);
  const pruneStale = () => {
    const generation = hooks.generation();
    rejected = rejected.filter((entry) => entry.generation === generation);
  };

  const follow = (work: Promise<unknown>) => {
    const settled = work.then(
      () => undefined,
      () => undefined
    );
    tail = settled;
    void settled.then(() => {
      if (tail !== settled) return;
      tail = null;
      failed = false;
    });
  };

  const enqueue = <T>(work: () => T | Promise<T>): Promise<T> => {
    if (tail) {
      const next = tail.then(work);
      follow(next);
      return next;
    }
    let result: T | Promise<T>;
    try {
      result = work();
    } catch (error) {
      failed = false;
      return Promise.reject(error);
    }
    if (isPromise(result)) {
      follow(result);
      return result;
    }
    failed = false;
    return Promise.resolve(result);
  };

  const fail = () => {
    failed = true;
    if (sealing) sealFailed = true;
  };

  /** Writes a draft once; later writes of the same draft share that outcome. */
  const writeOnce = (target: InputDraft): boolean => {
    if (written.has(target)) return true;
    if (!hooks.write(target)) return false;
    written.add(target);
    settleRejected(target);
    return true;
  };

  const blocked = () => {
    pruneStale();
    return rejected.length > 0;
  };

  /** Writes a finished draft after everything queued before it; a refused one stays rejected. */
  const queueWrite = (finished: InputDraft) => {
    void enqueue(() => {
      if (finished.generation !== hooks.generation() || writeOnce(finished)) return;
      failed = true;
      reject(finished);
      if (draft === null && hooks.restore(finished)) draft = finished;
    });
  };

  const writeSealed = (sealed: InputDraft | null, generation: number) => {
    if (hooks.generation() !== generation) throw new XlsxCommandAdmissionError('document-replaced');
    if (failed || blocked()) throw new XlsxCommandAdmissionError('input-failed');
    if (!sealed || sealed.generation !== generation) return;
    if (!writeOnce(sealed)) {
      failed = true;
      reject(sealed);
      throw new XlsxCommandAdmissionError('input-failed');
    }
    if (draft === sealed) {
      draft = null;
      hooks.close(sealed);
    }
  };

  return {
    get draft() {
      return draft;
    },
    get rejected() {
      return rejected;
    },
    get pending() {
      return tail !== null;
    },
    setDraft(next) {
      draft = next;
    },
    input(write) {
      return enqueue(() => {
        const settle = (ok: boolean | void) => {
          if (ok === false) fail();
        };
        let result: boolean | void | Promise<boolean | void>;
        try {
          result = write();
        } catch {
          result = false;
        }
        return isPromise(result) ? result.then(settle, () => settle(false)) : settle(result);
      });
    },
    submit(finished) {
      if (draft === finished) draft = null;
      if (!tail) {
        if (finished.generation !== hooks.generation()) return true;
        if (writeOnce(finished)) return true;
        reject(finished);
        draft = finished;
        return false;
      }
      queueWrite(finished);
      return true;
    },
    discard(target) {
      settleRejected(target);
    },
    settle() {
      const current = draft;
      if (!current) return true;
      if (current.generation !== hooks.generation()) {
        draft = null;
        return true;
      }
      if (tail) {
        draft = null;
        hooks.close(current);
        queueWrite(current);
        return true;
      }
      const known = rejected.some((entry) => sameCell(entry, current));
      if (!writeOnce(current)) {
        if (!known) return false;
        reject(current);
      }
      if (draft === current) {
        draft = null;
        hooks.close(current);
      }
      return true;
    },
    runAfterPendingInput(operation) {
      if (blocked()) return Promise.reject(new XlsxCommandAdmissionError('input-failed'));
      sealing = true;
      sealFailed = false;
      let seal: InputSeal;
      try {
        seal = hooks.seal();
      } finally {
        sealing = false;
      }
      if (seal.refused) return Promise.reject(new XlsxCommandAdmissionError(seal.refused));
      if (sealFailed) return Promise.reject(new XlsxCommandAdmissionError('input-failed'));
      const generation = hooks.generation();
      const composition = seal.composition;
      if (composition) {
        const composed = composition.then((ended) => {
          if (!ended) return undefined;
          hooks.sync();
          return draft;
        });
        return enqueue(async () => {
          const sealed = await composed;
          if (sealed === undefined) throw new XlsxCommandAdmissionError('input-failed');
          writeSealed(sealed, generation);
          return operation();
        });
      }
      hooks.sync();
      const sealed = draft;
      return enqueue(() => {
        writeSealed(sealed, generation);
        return operation();
      });
    },
  };
}
