/**
 * The resident worker's request scheduler. Messages are classified into lanes
 * and run one unit at a time, highest lane first. Before a unit that a later
 * message could overtake (a reorderable unit or a background slice), the
 * scheduler yields a turn of the event loop so messages that arrived meanwhile
 * are classified first; any other unit starts at once.
 *
 * Foreground messages keep their arrival order: a unit overtakes an earlier one
 * of another lane only when that one is `reorderable` (it detects staleness and
 * answers superseded itself) and its own lane ranks higher. Background tasks are
 * resumable and yield to every foreground message at slice boundaries.
 */

/** @internal */
export type SchedulerLane = 'input' | 'collab' | 'interactive' | 'background';

const FOREGROUND_LANES = ['input', 'collab', 'interactive'] as const;
type ForegroundLane = (typeof FOREGROUND_LANES)[number];
const RANK: Record<SchedulerLane, number> = { input: 0, collab: 1, interactive: 2, background: 3 };

/** Slice budget while the user is typing or deciding. */
export const INPUT_SLICE_MS = 8;
/** Slice budget otherwise. */
export const IDLE_SLICE_MS = 24;
/** How long after the last user input the input budget applies. */
export const INPUT_ACTIVE_MS = 1000;

/** @internal */
export interface SchedulerMessage {
  lane: ForegroundLane;
  run(): Promise<void> | void;
  /** User input: it selects the input slice budget. */
  userInput?: boolean;
  /**
   * Also restarts `idleAfterInputMs` holds. Leave it unset for input that can
   * stream without a pause and gets no reply, such as collaboration updates.
   */
  holdsIdleTasks?: boolean;
  /** A later unit of a higher lane may run first; this one then detects that itself. */
  reorderable?: boolean;
  /** Changes the document: bumps the scheduler version before it runs. */
  mutates?: boolean;
  /** Changes the presented frame or session: bumps the scheduler generation before it runs. */
  reframes?: boolean;
  /** Queued units this one replaces (latest wins). */
  key?: string;
  /** Replaced by a later unit with this `key` while still queued; `supersede` answers it. */
  replaceableBy?: string;
  supersede?(): void;
}

/** @internal */
export type TaskStep = 'done' | 'yield';

/** @internal A resumable background unit: each `run` does at most about one slice of work. */
export interface SchedulerTask {
  kind: string;
  /** Scheduler version and generation when the task was made. */
  version: number;
  generation: number;
  run(budgetMs: number): Promise<TaskStep> | TaskStep;
  /** Hold the task until the user has been idle this long (ms after the last user input). */
  idleAfterInputMs?: number;
  /**
   * The version or generation moved since the task was made: 'cancel' drops it
   * (after `cancel`), 'continue' runs it on with the current stamps.
   */
  onStale?(): 'cancel' | 'continue';
  /** Answers the task's request when it is dropped before it is done. */
  cancel?(): void;
  /** A run threw. */
  fail?(error: unknown): void;
}

/**
 * Work that is a pure function of serialisable input. An executor may run
 * `compute` elsewhere (a helper worker looks it up by `kind`); `install` always
 * runs here, as a background unit, and only while the stamps still match.
 * @internal
 */
export interface PureTask<Input = unknown, Result = unknown> {
  kind: string;
  version: number;
  generation: number;
  input: Input;
  transfer?: Transferable[];
  compute(input: Input): Result;
  install(result: Result): Promise<void> | void;
  /**
   * The result went stale before install, install failed, or the executor
   * failed after `transfer` detached the input (an executor failure otherwise
   * computes here).
   */
  cancel?(reason?: unknown): void;
}

/** @internal Runs pure tasks' `compute` somewhere else. */
export interface TaskExecutor {
  run<Input, Result>(job: { kind: string; input: Input; transfer?: Transferable[] }): Promise<Result>;
}

/** @internal */
export interface SchedulerHost {
  now(): number;
  /** Calls back after a turn of the event loop. */
  turn(callback: () => void): void;
  /** Calls back after `ms`; returns a cancel. */
  timer(callback: () => void, ms: number): () => void;
  /** Absent: pure tasks compute here, in background slices. */
  executor?: TaskExecutor;
  /** A foreground unit threw past its own handling. */
  failed?(error: unknown): void;
}

interface QueuedMessage extends SchedulerMessage {
  seq: number;
}

/** @internal */
export interface ResidentScheduler {
  submit(message: SchedulerMessage): void;
  schedule(task: SchedulerTask): void;
  dispatch<Input, Result>(task: PureTask<Input, Result>): void;
  /** The slice budget for the next background unit. */
  budget(): number;
  readonly version: number;
  readonly generation: number;
  /** Bumps the stamps outside a message (a background unit that changes the document). */
  bump(changes: { version?: boolean; generation?: boolean }): void;
  /** Queued foreground units and background tasks. */
  pending(): { foreground: number; background: number };
}

/** @internal */
export function createResidentScheduler(host: SchedulerHost): ResidentScheduler {
  const lanes: Record<ForegroundLane, QueuedMessage[]> = { input: [], collab: [], interactive: [] };
  let background: SchedulerTask[] = [];
  let seq = 0;
  let version = 0;
  let generation = 0;
  let lastInputAt = Number.NEGATIVE_INFINITY;
  let lastHoldAt = Number.NEGATIVE_INFINITY;
  let running = false;
  let pumpQueued = false;
  let cancelTimer: (() => void) | null = null;

  const budget = (): number =>
    host.now() - lastInputAt < INPUT_ACTIVE_MS ? INPUT_SLICE_MS : IDLE_SLICE_MS;

  function wake(now = false): void {
    if (running || pumpQueued) return;
    pumpQueued = true;
    (now ? queueMicrotask : host.turn)(() => {
      pumpQueued = false;
      void pump();
    });
  }

  /** Whether `candidate` must wait for an earlier unit of another lane. */
  function blocked(candidate: QueuedMessage): boolean {
    for (const lane of FOREGROUND_LANES) {
      if (lane === candidate.lane) continue;
      for (const queued of lanes[lane]) {
        if (queued.seq > candidate.seq) break;
        if (!(queued.reorderable && RANK[candidate.lane] < RANK[queued.lane])) return true;
      }
    }
    return false;
  }

  function nextForeground(): QueuedMessage | null {
    for (const lane of FOREGROUND_LANES) {
      const head = lanes[lane][0];
      if (head && !blocked(head)) return head;
    }
    return null;
  }

  /** The next foreground units to run as one: a collab head takes the collab units right behind it. */
  function pickForeground(): QueuedMessage[] | null {
    const head = nextForeground();
    if (!head) return null;
    const units = [lanes[head.lane].shift()!];
    if (head.lane === 'collab') {
      while (lanes.collab[0] && !blocked(lanes.collab[0])) units.push(lanes.collab.shift()!);
    }
    return units;
  }

  function readyAt(task: SchedulerTask): number {
    return task.idleAfterInputMs === undefined ? Number.NEGATIVE_INFINITY : lastHoldAt + task.idleAfterInputMs;
  }

  async function runForeground(units: QueuedMessage[]): Promise<void> {
    for (const unit of units) {
      if (unit.mutates) version += 1;
      if (unit.mutates || unit.reframes) generation += 1;
      try {
        await unit.run();
      } catch (error) {
        host.failed?.(error);
      }
    }
  }

  /** Runs one slice of the first ready task; false when none is ready. */
  async function runBackground(): Promise<boolean> {
    const now = host.now();
    const index = background.findIndex((task) => readyAt(task) <= now);
    if (index < 0) return false;
    const [task] = background.splice(index, 1);
    try {
      if (task.version !== version || task.generation !== generation) {
        if ((task.onStale?.() ?? 'cancel') === 'cancel') {
          task.cancel?.();
          return true;
        }
        task.version = version;
        task.generation = generation;
      }
      if ((await task.run(budget())) === 'yield') background.push(task);
    } catch (error) {
      if (task.fail) task.fail(error);
      else host.failed?.(error);
    }
    return true;
  }

  function armTimer(): void {
    cancelTimer?.();
    cancelTimer = null;
    if (background.length === 0) return;
    const at = Math.min(...background.map(readyAt));
    cancelTimer = host.timer(() => {
      cancelTimer = null;
      wake();
    }, Math.max(1, at - host.now()));
  }

  async function pump(): Promise<void> {
    if (running) return;
    running = true;
    let ran = false;
    try {
      const units = pickForeground();
      if (units) {
        ran = true;
        await runForeground(units);
      } else if (background.length > 0) {
        ran = await runBackground();
      }
    } catch (error) {
      ran = true;
      host.failed?.(error);
    } finally {
      running = false;
    }
    // A later message cannot overtake a unit that is not reorderable.
    const next = ran ? nextForeground() : null;
    if (next) wake(!next.reorderable);
    else if (ran && background.length > 0) wake();
    else armTimer();
  }

  function settle(task: PureTask<unknown, unknown>, result: unknown): void {
    schedule({
      kind: `${task.kind}:install`,
      version: task.version,
      generation: task.generation,
      onStale: () => 'cancel',
      cancel: () => task.cancel?.(),
      fail: (error) => task.cancel?.(error),
      run: async (): Promise<TaskStep> => {
        await task.install(result);
        return 'done';
      },
    });
  }

  function schedule(task: SchedulerTask): void {
    background.push(task);
    if (!running) {
      cancelTimer?.();
      cancelTimer = null;
    }
    wake();
  }

  return {
    submit(message) {
      if (message.userInput) lastInputAt = host.now();
      if (message.holdsIdleTasks) lastHoldAt = host.now();
      if (message.key !== undefined) {
        const lane = lanes[message.lane];
        for (let index = lane.length - 1; index >= 0; index -= 1) {
          if (lane[index]!.replaceableBy !== message.key) continue;
          try {
            lane.splice(index, 1)[0]!.supersede?.();
          } catch (error) {
            host.failed?.(error);
          }
        }
      }
      lanes[message.lane].push({ ...message, seq: seq++ });
      cancelTimer?.();
      cancelTimer = null;
      wake(true);
    },
    schedule,
    dispatch<Input, Result>(task: PureTask<Input, Result>) {
      const pure = task as unknown as PureTask<unknown, unknown>;
      const local = (): void =>
        schedule({
          kind: pure.kind,
          version: pure.version,
          generation: pure.generation,
          onStale: () => 'cancel',
          cancel: () => pure.cancel?.(),
          fail: (error) => pure.cancel?.(error),
          run: async (): Promise<TaskStep> => {
            await pure.install(pure.compute(pure.input));
            return 'done';
          },
        });
      if (!host.executor) {
        local();
        return;
      }
      const transfers = (pure.transfer?.length ?? 0) > 0;
      host.executor
        .run<unknown, unknown>({ kind: pure.kind, input: pure.input, transfer: pure.transfer })
        .then(
          (result) => settle(pure, result),
          (error) => (transfers ? pure.cancel?.(error) : local())
        );
    },
    budget,
    get version() {
      return version;
    },
    get generation() {
      return generation;
    },
    bump(changes) {
      if (changes.version) version += 1;
      if (changes.version || changes.generation) generation += 1;
    },
    pending() {
      return {
        foreground: lanes.input.length + lanes.collab.length + lanes.interactive.length,
        background: background.length,
      };
    },
  };
}
