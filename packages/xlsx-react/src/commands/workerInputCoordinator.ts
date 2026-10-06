import { XlsxCommandAdmissionError } from './createXlsxCommandStore';
import type { InputCoordinatorHooks, InputDraft, InputSeal } from './inputCoordinator';

export type WorkerInputKind =
  | 'commit'
  | 'input'
  | 'clipboard'
  | 'chart'
  | 'command'
  | 'host'
  | 'plugin';

export interface WorkerInputTarget {
  readonly sheet: number;
  readonly target: string;
}

export interface WorkerInputIntent<T = unknown> extends WorkerInputTarget {
  readonly id: number;
  readonly generation: number;
  readonly kind: WorkerInputKind;
  readonly input: T | undefined | Promise<T | undefined>;
  readonly draft: InputDraft | null;
}

export interface WorkerInputOperationOptions<T = unknown> {
  readonly kind?: Exclude<WorkerInputKind, 'commit'>;
  readonly target?: WorkerInputTarget;
  readonly input?: T;
  readonly recover?: boolean;
  readonly barrier?: boolean;
  readonly prepare?: () => void | Promise<void>;
}

export interface WorkerInputLease {
  (): void;
  check(): void;
  refuse(error: unknown, barrier?: boolean): void;
}

export class WorkerInputRefusal extends XlsxCommandAdmissionError {
  constructor(message: string, readonly barrier = false) {
    super('input-failed');
    this.message = message;
  }
}

export function inputRefusal(error: unknown): WorkerInputRefusal | null {
  if (error instanceof WorkerInputRefusal) return error;
  if (error instanceof Error && (error.name === 'NotAllowedError' || error.name === 'NotFoundError' ||
    error instanceof RangeError || /32,?767|out of range|invalid (sheet|cell|input|formula)|(text|formula).*limit/i.test(error.message))) {
    return new WorkerInputRefusal(error.message);
  }
  return null;
}

export type WorkerInputOperation<T, R> = (
  intent: WorkerInputIntent<T>,
  markApplied: WorkerInputLease
) => R | Promise<R>;

export interface WorkerInputCoordinatorHooks
  extends Pick<InputCoordinatorHooks, 'generation' | 'seal' | 'sync'> {
  write(draft: InputDraft, markApplied: WorkerInputLease): boolean;
  capture(): WorkerInputTarget;
  isReady(): boolean;
  isAcknowledged?(): boolean;
  /** Waits passively for the edit peer. */
  whenReady(): Promise<void>;
  resolveDraft?(draft: InputDraft): Promise<InputDraft>;
  sameDraft?(a: InputDraft, b: InputDraft): boolean;
  /** Resolves after the pending text has crossed a browser paint boundary. */
  preview(draft: InputDraft): Promise<void>;
  requestHydration(reason: string): void | Promise<void>;
  flushEdits(): Promise<void>;
  acknowledgeEdits?(): Promise<void>;
  onError?(error: unknown): void;
  onRefusal?(error: unknown, draft?: InputDraft): void;
}

export class WorkerInputNotReadyError extends Error {
  constructor() {
    super('The edit peer and accepted input must be flushed before a synchronous mutation');
    this.name = 'WorkerInputNotReadyError';
  }
}

export interface WorkerInputCoordinator {
  readonly draft: InputDraft | null;
  readonly committed: readonly InputDraft[];
  readonly rejected: readonly InputDraft[];
  readonly unapplied: readonly WorkerInputIntent[];
  readonly pending: boolean;
  readonly error: unknown;
  setDraft(draft: InputDraft | null): void;
  /** Accepts a commit; its write waits for readiness and paint. */
  submit(draft: InputDraft): boolean;
  submitAsync(draft: InputDraft): Promise<void>;
  settle(): boolean;
  settleAsync(): Promise<void>;
  input(
    write: WorkerInputOperation<unknown, boolean | void>,
    options?: WorkerInputOperationOptions
  ): Promise<void>;
  /** Captures clipboard access in the gesture; false refuses its write. */
  clipboard<T, R>(
    capture: () => T | Promise<T>,
    operation: (input: T, intent: WorkerInputIntent<T>, markApplied: WorkerInputLease) => R | Promise<R>,
    target?: WorkerInputTarget,
    prepare?: (input: T) => void | Promise<void>,
    recover?: boolean
  ): Promise<R>;
  /** Async mutators call markApplied before awaiting acknowledgement of an applied edit. */
  runAfterPendingInput<T, R>(
    operation: WorkerInputOperation<T, R>,
    options?: WorkerInputOperationOptions<T>
  ): Promise<R>;
  runSync<T, R>(
    operation: (intent: WorkerInputIntent<T>, markApplied: WorkerInputLease) => R,
    options?: WorkerInputOperationOptions<T>
  ): R;
  requestHydration(reason: string): Promise<void>;
  flush(): Promise<void>;
  fail(error: unknown): void;
  recover(): Promise<void>;
  drain(): Promise<void>;
  retire(): void;
  reset(): void;
}

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let rejectPromise!: (error: unknown) => void;
  let rejection: { error: unknown } | null = null;
  const cancellations = new Set<(error: unknown) => void>();
  const promise = new Promise<T>((done, failed) => {
    resolve = done;
    rejectPromise = failed;
  });
  void promise.catch(() => {});
  return {
    promise, resolve,
    reject(error: unknown) {
      rejection = { error };
      rejectPromise(error);
      for (const cancel of cancellations) cancel(error);
      cancellations.clear();
    },
    wait<R>(value: R | Promise<R>): Promise<R> {
      if (rejection) {
        void Promise.resolve(value).catch(() => {});
        return Promise.reject(rejection.error);
      }
      return new Promise<R>((done, failed) => {
        const cancel = (error: unknown) => { cancellations.delete(cancel); failed(error); };
        cancellations.add(cancel);
        void Promise.resolve(value).then((result) => {
          cancellations.delete(cancel);
          done(result);
        }, cancel);
      });
    },
  };
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

function isPromise<T>(value: T | Promise<T>): value is Promise<T> {
  return typeof (value as Promise<T> | null)?.then === 'function';
}

interface Entry {
  intent: WorkerInputIntent;
  state: 'queued' | 'running' | 'failed';
  applied: boolean;
  draftWrite: boolean;
  recover: boolean;
  refused: boolean;
  prepared?: Promise<void>;
  barrier: boolean;
  barrierAfter: number;
  error?: unknown;
  run(markApplied: WorkerInputLease): unknown | Promise<unknown>;
  resolve(value: unknown): void;
  reject(error: unknown): void;
}

/** One FIFO for worker input, with retained unapplied work after failure. */
export function createWorkerInputCoordinator(
  hooks: WorkerInputCoordinatorHooks
): WorkerInputCoordinator {
  let draft: InputDraft | null = null;
  let rejected: readonly InputDraft[] = [];
  let entries: Entry[] = [];
  let snapshots = new WeakMap<InputDraft, InputDraft>();
  let written = new WeakSet<InputDraft>();
  let preparedDrafts = new WeakMap<InputDraft, Promise<InputDraft>>();
  let preparation: Promise<void> = Promise.resolve();
  let generation = hooks.generation();
  let epoch = 0;
  let cycle = deferred<never>();
  let replaced = deferred<never>();
  let active: typeof cycle | null = null;
  let work: Promise<void> | null = null;
  let recovery: Promise<void> | null = null;
  let failure: { error: unknown } | null = null;
  let reported = false;
  let inputFailed = false;
  let syncRunning = false;
  let nextId = 1;
  let refused: { id: number; error: unknown }[] = [];
  let acknowledgedRefusal = 0;

  const snapshot = (value: InputDraft): InputDraft => {
    let saved = snapshots.get(value);
    if (!saved) {
      saved = Object.freeze({ ...value });
      snapshots.set(value, saved);
      snapshots.set(saved, saved);
    }
    return saved;
  };

  const reset = () => {
    const error = new XlsxCommandAdmissionError('document-replaced');
    cycle.reject(error);
    replaced.reject(error);
    for (const entry of entries) entry.reject(error);
    cycle = deferred<never>();
    replaced = deferred<never>();
    epoch += 1;
    active = null;
    work = null;
    recovery = null;
    entries = [];
    refused = [];
    acknowledgedRefusal = 0;
    rejected = [];
    draft = null;
    snapshots = new WeakMap();
    written = new WeakSet();
    preparedDrafts = new WeakMap();
    preparation = Promise.resolve();
    failure = null;
    reported = false;
    inputFailed = false;
    generation = hooks.generation();
  };

  const current = () => {
    if (generation !== hooks.generation()) reset();
  };

  const check = (intent: WorkerInputIntent, lease: typeof cycle) => {
    if (lease !== cycle || intent.generation !== hooks.generation()) {
      throw new XlsxCommandAdmissionError('document-replaced');
    }
    if (failure) throw failure.error;
  };

  const wait = <T>(promise: T | Promise<T>, lease: typeof cycle): Promise<T> =>
    lease.wait(promise);

  const settleRejected = (value: InputDraft) => {
    rejected = rejected.filter((entry) => !sameCell(entry, value));
  };

  const coordinatorError = (): { error: unknown } | null => failure;

  const fail = (error: unknown) => {
    if (failure) return;
    failure = { error };
    cycle.reject(error);
    for (const entry of entries) entry.reject(error);
    for (const entry of entries) {
      if (!entry.applied && !entry.recover) {
        hooks.onRefusal?.(error);
        if (entry.intent.draft && !written.has(entry.intent.draft)) {
          entry.intent = Object.freeze({ ...entry.intent, kind: 'commit' });
          entry.recover = true;
          entry.run = (applied) => write(entry.intent.draft, entry.intent, cycle, applied);
        }
      }
    }
    entries = entries.filter((entry) => !entry.applied && entry.recover);
    if (!reported) {
      reported = true;
      try {
        hooks.onError?.(error);
      } catch {}
    }
  };

  const prepareDraft = (value: InputDraft): Promise<InputDraft> => {
    const existing = preparedDrafts.get(value);
    if (existing) return existing;
    const prepared = Promise.resolve().then(async () => {
      const resolved = hooks.resolveDraft ? await hooks.resolveDraft(value) : value;
      preparedDrafts.set(resolved, prepared);
      await hooks.preview(resolved);
      return resolved;
    });
    preparedDrafts.set(value, prepared);
    void prepared.catch(() => {});
    return prepared;
  };

  const write = async (value: InputDraft | null, intent: WorkerInputIntent, lease: typeof cycle, markApplied: WorkerInputLease) => {
    if (!value || written.has(value)) return;
    check(intent, lease);
    if (value.generation !== intent.generation) {
      throw new XlsxCommandAdmissionError('document-replaced');
    }
    const resolved = await wait(prepareDraft(value), lease);
    check(intent, lease);
    if (!hooks.isReady()) throw new WorkerInputNotReadyError();
    if (!hooks.write(resolved, markApplied)) {
      inputFailed = true;
      const correction = entries.some(
        (entry) =>
          entry.state === 'queued' &&
          entry.intent.kind === 'commit' &&
          entry.intent.draft !== null &&
          sameCell(entry.intent.draft, value)
      );
      if (!correction) {
        rejected = [...rejected.filter((entry) => !sameCell(entry, value)), value];
        if (draft === null) draft = value;
      }
      throw new XlsxCommandAdmissionError('input-failed');
    }
    written.add(value);
    written.add(resolved);
    const committed = entries.find((entry) => entry.intent.id === intent.id);
    if (committed?.intent.kind === 'commit') committed.applied = true;
    settleRejected(value);
    entries = entries.filter((entry) => {
      if (entry.state !== 'failed' || !entry.intent.draft || !sameCell(entry.intent.draft, value)) {
        return true;
      }
      written.add(entry.intent.draft);
      if (entry.intent.kind === 'commit') entry.applied = true;
      return entry.intent.kind !== 'commit';
    });
    if (draft && snapshot(draft) === value) draft = null;
  };

  const pump = (recovering = false): Promise<void> => {
    if (failure || (recovery && !recovering) || syncRunning) return Promise.resolve();
    if (active === cycle && work) return work;
    if (!entries.some((entry) => entry.state === 'queued')) return Promise.resolve();
    const lease = cycle;
    active = lease;
    const draining = async () => {
      while (lease === cycle && !failure) {
        let entry = entries[0];
        if (!entry) break;
        if (entry.state === 'failed' && entry.draftWrite && entry.intent.kind !== 'clipboard' &&
          entry.intent.draft && !written.has(entry.intent.draft)) {
          const blockedDraft = entry.intent.draft;
          const correction = entries[1];
          if (correction?.state === 'queued' && correction.intent.kind === 'commit' &&
            correction.intent.draft && sameCell(correction.intent.draft, blockedDraft)) {
            entries = [correction, ...entries.filter((candidate) => candidate !== correction)];
            entry = correction;
          }
        }
        if (entry.state !== 'queued') {
          for (const blocked of entries) {
            if (blocked.state === 'queued') blocked.reject(entry.error ?? new XlsxCommandAdmissionError('input-failed'));
          }
          break;
        }
        entry.state = 'running';
        try {
          check(entry.intent, lease);
          if (entry.prepared) await wait(entry.prepared, lease);
          await wait(hooks.whenReady(), lease);
          check(entry.intent, lease);
          if (!hooks.isReady()) throw new WorkerInputNotReadyError();
          if (!recovering && hooks.acknowledgeEdits) await wait(hooks.acknowledgeEdits(), lease);
          const markApplied: WorkerInputLease = Object.assign(() => {
            check(entry.intent, lease);
            entry.applied = true;
          }, {
            check: () => check(entry.intent, lease),
            refuse(error: unknown, barrier = false) {
              check(entry.intent, lease);
              entry.refused = true;
              if (barrier) refused.push({ id: entry.intent.id, error });
              hooks.onRefusal?.(error, entry.intent.kind === 'commit' ? entry.intent.draft ?? undefined : undefined);
            },
          });
          if (entry.barrier) {
            const covering = refused.find((item) => item.id > entry.barrierAfter && item.id < entry.intent.id);
            if (covering) {
              acknowledgedRefusal = Math.max(acknowledgedRefusal, covering.id);
              throw new WorkerInputRefusal(covering.error instanceof Error ? covering.error.message : String(covering.error));
            }
          }
          const result = await wait(entry.run(markApplied), lease);
          if (!recovering && hooks.acknowledgeEdits) await wait(hooks.acknowledgeEdits(), lease);
          check(entry.intent, lease);
          if (!entry.refused) entry.applied = true;
          entry.resolve(result);
          entries = entries.filter((candidate) => candidate !== entry);
        } catch (error) {
          if (lease !== cycle) return;
          entry.reject(error);
          const local = inputRefusal(error);
          if (local) {
            if (local.barrier) refused.push({ id: entry.intent.id, error: local });
            hooks.onRefusal?.(local, entry.intent.kind === 'commit' ? entry.intent.draft ?? undefined : undefined);
            entries = entries.filter((candidate) => candidate !== entry);
            continue;
          }
          if (entry.applied) entries = entries.filter((candidate) => candidate !== entry);
          else {
            entry.state = 'failed';
            entry.error = error;
          }
          if (error instanceof XlsxCommandAdmissionError) {
            if (error.code === 'document-replaced') {
              hooks.onRefusal?.(error);
              entries = entries.filter((candidate) => candidate !== entry);
              continue;
            }
            if (error.code === 'input-failed') {
              inputFailed = true;
              hooks.onRefusal?.(error);
            }
          } else {
            fail(error);
            return;
          }
        }
      }
      if (lease === cycle) inputFailed = entries.some((entry) => entry.state === 'failed');
    };
    const running = draining().finally(() => {
      if (active === lease) {
        active = null;
        work = null;
        if (entries[0]?.state === 'queued') void pump();
      }
    });
    work = running;
    return running;
  };

  const intent = <T>(
    kind: WorkerInputKind,
    input: WorkerInputIntent<T>['input'],
    sealed: InputDraft | null,
    target = hooks.capture()
  ): WorkerInputIntent<T> =>
    Object.freeze({
      id: nextId++,
      generation,
      kind,
      sheet: target.sheet,
      target: target.target,
      input,
      draft: sealed,
    });

  const record = <R>(
    saved: WorkerInputIntent,
    run: (entry: Entry, lease: typeof cycle, markApplied: WorkerInputLease) => R | Promise<R>
  ) => {
    const result = deferred<R>();
    const entry: Entry = {
      intent: saved,
      state: 'queued',
      applied: false,
      draftWrite: false,
      recover: true,
      refused: false,
      barrier: false,
      barrierAfter: acknowledgedRefusal,
      run: (markApplied) => run(entry, cycle, markApplied),
      resolve: (value) => result.resolve(value as R),
      reject: result.reject,
    };
    entries.push(entry);
    return { entry, result };
  };

  const accept = <R>(
    saved: WorkerInputIntent,
    run: (entry: Entry, lease: typeof cycle, markApplied: WorkerInputLease) => R | Promise<R>,
    draftWrite = false,
    options: WorkerInputOperationOptions = {}
  ): Promise<R> => {
    const { entry, result } = record(saved, run);
    entry.draftWrite = draftWrite;
    entry.recover = options.recover ?? true;
    entry.barrier = options.barrier ?? false;
    if (options.prepare) {
      entry.prepared = preparation.then(options.prepare);
      preparation = entry.prepared.catch(() => {});
    }
    if (failure) entry.reject(failure.error);
    const blocked = entries[0];
    if (blocked?.state === 'failed' && !(saved.kind === 'commit' && saved.draft &&
      blocked.draftWrite && blocked.intent.kind !== 'clipboard' && blocked.intent.draft &&
      !written.has(blocked.intent.draft) &&
      sameCell(saved.draft, blocked.intent.draft))) {
      entry.reject(blocked.error ?? new XlsxCommandAdmissionError('input-failed'));
    }
    void pump();
    return result.promise;
  };

  const submit = (finished: InputDraft): Promise<void> => {
    current();
    const saved = snapshot(finished);
    if (saved.generation !== generation) {
      return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    }
    if (draft === finished || (draft && snapshot(draft) === saved)) draft = null;
    settleRejected(saved);
    const accepted = intent('commit', saved, saved, {
      sheet: saved.sheet,
      target: `cell:${saved.row}:${saved.col}`,
    });
    return accept(accepted, (entry, lease, applied) => write(entry.intent.draft, entry.intent, lease, applied), true, {
      prepare: async () => {
        const resolved = await prepareDraft(saved);
        const entry = entries.find((item) => item.intent.id === accepted.id);
        if (entry) entry.intent = Object.freeze({ ...entry.intent, draft: resolved });
      },
    });
  };

  const runAfterPendingInput = <T, R>(
    operation: WorkerInputOperation<T, R>,
    options: WorkerInputOperationOptions<T> = {},
    captured?: { input: T | Promise<T> }
  ): Promise<R> => {
    current();
    if (rejected.length > 0) {
      return Promise.reject(new XlsxCommandAdmissionError('input-failed'));
    }
    let seal: InputSeal;
    try {
      seal = hooks.seal();
    } catch (error) {
      return Promise.reject(error);
    }
    if (seal.refused) return Promise.reject(new XlsxCommandAdmissionError(seal.refused));
    if (!seal.composition) hooks.sync();
    const original = draft;
    const saved = intent(
      options.kind ?? 'command',
      captured ? captured.input : structuredClone(options.input),
      original ? snapshot(original) : null,
      options.target
    );
    const acceptedEpoch = epoch;
    const composed = seal.composition?.then((ended) => {
      if (acceptedEpoch !== epoch || saved.generation !== hooks.generation()) {
        throw new XlsxCommandAdmissionError('document-replaced');
      }
      if (!ended) throw new XlsxCommandAdmissionError('input-failed');
      let composedDraft = saved.draft;
      if (saved.draft && draft && sameCell(saved.draft, draft) && (hooks.sameDraft?.(saved.draft, draft) ?? true)) {
        hooks.sync();
        composedDraft = draft ? snapshot(draft) : null;
      } else {
        const committed = entries.filter(
          (entry) =>
            entry.intent.id > saved.id &&
            entry.intent.kind === 'commit' &&
            saved.draft !== null &&
            entry.intent.draft !== null &&
            sameCell(saved.draft, entry.intent.draft) &&
            (hooks.sameDraft?.(saved.draft, entry.intent.draft) ?? true)
        );
        composedDraft = committed[committed.length - 1]?.intent.draft ?? saved.draft;
      }
      const entry = entries.find((candidate) => candidate.intent.id === saved.id);
      if (entry) entry.intent = Object.freeze({ ...entry.intent, draft: composedDraft });
      return composedDraft;
    });
    void composed?.catch(() => {});
    return accept(saved, async (entry, activeLease, markApplied) => {
      const sealed = composed ? await wait(composed, activeLease) : entry.intent.draft;
      check(entry.intent, activeLease);
      if (inputFailed || rejected.length > 0) {
        throw new XlsxCommandAdmissionError('input-failed');
      }
      if (sealed !== entry.intent.draft) {
        entry.intent = Object.freeze({ ...entry.intent, draft: sealed });
      }
      if (captured) {
        const input = await wait(captured.input, activeLease);
        check(entry.intent, activeLease);
        entry.intent = Object.freeze({ ...entry.intent, input });
      }
      await write(sealed, entry.intent, activeLease, Object.assign(() => {}, { check: markApplied.check, refuse: markApplied.refuse }));
      if (sealed && hooks.acknowledgeEdits) await wait(hooks.acknowledgeEdits(), activeLease);
      check(entry.intent, activeLease);
      const result = operation(entry.intent as WorkerInputIntent<T>, markApplied);
      if (!isPromise(result) && result !== false && !entry.refused) markApplied();
      return result;
    }, true, { ...options, prepare: async () => {
      if (saved.draft && !composed) {
        const resolved = await prepareDraft(saved.draft);
        const entry = entries.find((item) => item.intent.id === saved.id);
        if (entry) entry.intent = Object.freeze({ ...entry.intent, draft: resolved });
      }
      await options.prepare?.();
    } });
  };

  const requestHydration = async (reason: string) => {
    current();
    const acceptedEpoch = epoch;
    try {
      await replaced.wait(Promise.resolve(hooks.requestHydration(reason)));
      if (acceptedEpoch !== epoch || generation !== hooks.generation()) {
        throw new XlsxCommandAdmissionError('document-replaced');
      }
    } catch (error) {
      if (acceptedEpoch === epoch && generation === hooks.generation()) fail(error);
      throw error;
    }
  };

  const assertReady = () => {
    if (failure) throw failure.error;
    if (!hooks.isReady() || hooks.isAcknowledged?.() === false || entries.length > 0 || recovery || syncRunning || draft) {
      throw new WorkerInputNotReadyError();
    }
  };

  return {
    get draft() {
      return draft;
    },
    get committed() {
      return entries.flatMap((entry) => {
        const value = entry.intent.draft;
        return entry.draftWrite && !entry.applied && value && !written.has(value) ? [value] : [];
      });
    },
    get rejected() {
      return rejected;
    },
    get unapplied() {
      return entries.filter((entry) => !entry.applied).map((entry) => entry.intent);
    },
    get pending() {
      return entries.length > 0 || active !== null || recovery !== null;
    },
    get error() {
      return failure?.error ?? null;
    },
    setDraft(next) {
      current();
      draft = next;
    },
    submit(finished) {
      if (!finished.value.startsWith('=') && [...finished.value.replace(/^'/, '')].length > 32767) {
        hooks.onRefusal?.(new WorkerInputRefusal('Cell text exceeds 32,767 characters'), finished);
        return false;
      }
      void submit(finished).catch(() => {});
      return true;
    },
    submitAsync: submit,
    settle() {
      if (draft) void submit(draft).catch(() => {});
      return true;
    },
    settleAsync() {
      return draft ? submit(draft) : Promise.resolve();
    },
    input(operation, options = {}) {
      current();
      const saved = intent(
        options.kind ?? 'input',
        structuredClone(options.input),
        draft ? snapshot(draft) : null,
        options.target
      );
      return accept(saved, async (entry, lease, markApplied) => {
        const writing = operation(entry.intent, markApplied);
        if (!isPromise(writing) && writing !== false) markApplied();
        const result = await wait(writing, lease);
        if (result === false) throw new XlsxCommandAdmissionError('input-failed');
      });
    },
    clipboard(capture, operation, target, prepare, recover = true) {
      current();
      const origin = { ...(target ?? hooks.capture()) };
      const acceptedEpoch = epoch;
      let data: ReturnType<typeof capture>;
      try {
        const captured = capture();
        data = isPromise(captured)
          ? captured.then((value) => structuredClone(value))
          : structuredClone(captured);
      } catch (error) {
        return Promise.reject(error);
      }
      if (acceptedEpoch !== epoch || generation !== hooks.generation()) {
        if (isPromise(data)) void data.catch(() => {});
        return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
      }
      if (isPromise(data)) void data.catch(() => {});
      return runAfterPendingInput(
        async (intent, markApplied) => {
          markApplied.check();
          const writing = operation(intent.input as Awaited<typeof data>, intent, markApplied);
          if (!isPromise(writing) && writing !== false) markApplied();
          const result = await writing;
          if (result === false) throw new XlsxCommandAdmissionError('input-failed');
          return result;
        },
        { kind: 'clipboard', target: origin, recover, prepare: prepare ? async () => prepare(await data) : undefined },
        { input: data }
      );
    },
    runAfterPendingInput,
    runSync<T, R>(
      operation: (intent: WorkerInputIntent<T>, markApplied: WorkerInputLease) => R,
      options: WorkerInputOperationOptions<T> = {}
    ) {
      current();
      assertReady();
      const seal = hooks.seal();
      if (seal.refused) throw new XlsxCommandAdmissionError(seal.refused);
      if (seal.composition) throw new WorkerInputNotReadyError();
      hooks.sync();
      assertReady();
      const saved = intent<T>(
        options.kind ?? 'host',
        structuredClone(options.input),
        null,
        options.target
      );
      syncRunning = true;
      const { entry, result } = record(saved, (currentEntry, lease, markApplied) => {
        check(currentEntry.intent, lease);
        return operation(saved, markApplied);
      });
      entry.state = 'running';
      try {
        const lease = cycle;
        const value = operation(saved, Object.assign(() => {
          check(saved, lease);
          entry.applied = true;
        }, { check: () => check(saved, lease), refuse(error: unknown, barrier = false) {
          entry.refused = true;
          if (barrier) refused.push({ id: saved.id, error });
          hooks.onRefusal?.(error);
        } }));
        entry.applied = true;
        entries = entries.filter((candidate) => candidate !== entry);
        result.resolve(value);
        return value;
      } catch (error) {
        entry.state = 'failed';
        entry.error = error;
        entry.reject(error);
        if (entry.applied) entries = entries.filter((candidate) => candidate !== entry);
        const local = inputRefusal(error);
        if (local) {
          hooks.onRefusal?.(local);
          entries = entries.filter((candidate) => candidate !== entry);
        } else if (error instanceof XlsxCommandAdmissionError) {
          if (error.code === 'input-failed') inputFailed = true;
        } else {
          fail(error);
        }
        throw error;
      } finally {
        syncRunning = false;
        void pump();
      }
    },
    requestHydration,
    flush() {
      current();
      if (failure) return Promise.reject(failure.error);
      const lease = cycle;
      const barrierAfter = acknowledgedRefusal;
      const flushed = runAfterPendingInput(async (_, markApplied) => {
        markApplied();
        await hooks.flushEdits();
      }, { kind: 'host', recover: false, barrier: true });
      void flushed.catch(() => {});
      return Promise.all([requestHydration('flush'), flushed]).then(async () => {
        const throwFailure = () => { if (failure) throw failure.error; };
        let watermark: number;
        do {
          await wait(pump(), lease);
          current();
          if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
          throwFailure();
          const covering = refused.find((item) => item.id > barrierAfter);
          if (covering) {
            acknowledgedRefusal = Math.max(acknowledgedRefusal, covering.id);
            throw new WorkerInputRefusal(covering.error instanceof Error ? covering.error.message : String(covering.error));
          }
          if (entries.some((entry) => !entry.applied)) {
            throw new XlsxCommandAdmissionError('input-failed');
          }
          watermark = nextId;
          await wait(hooks.flushEdits(), lease);
          current();
          if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
          throwFailure();
        } while (nextId !== watermark || entries.length > 0);
      });
    },
    fail,
    recover() {
      current();
      if (recovery) return recovery;
      const lease = cycle;
      const recovered = async () => {
        await requestHydration('recovery');
        if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
        if (work) await work;
        if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
        cycle = deferred<never>();
        failure = null;
        inputFailed = false;
        rejected = [];
        preparedDrafts = new WeakMap();
        preparation = Promise.resolve();
        for (const entry of entries) { entry.state = 'queued'; entry.prepared = undefined; }
        do {
          await pump(true);
        } while (entries[0]?.state === 'queued' && !coordinatorError());
        const error = coordinatorError();
        if (error) throw error.error;
        if (entries.length > 0) throw new XlsxCommandAdmissionError('input-failed');
      };
      const result = Promise.resolve().then(recovered).finally(() => {
        if (recovery === result) {
          recovery = null;
          void pump();
        }
      });
      recovery = result;
      return result;
    },
    async drain() {
      if (draft) await submit(draft);
      await pump();
      if (entries.some((entry) => !entry.applied)) {
        const error = failure?.error ?? new XlsxCommandAdmissionError('input-failed');
        hooks.onRefusal?.(error);
        throw error;
      }
      await hooks.flushEdits();
    },
    retire() {
      const error = new XlsxCommandAdmissionError('document-replaced');
      for (const entry of entries) {
        if (entry.intent.kind !== 'clipboard') entry.reject(error);
      }
    },
    reset,
  };
}
