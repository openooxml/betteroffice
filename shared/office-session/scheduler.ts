export type SchedulerLane = 'input' | 'collab' | 'interactive' | 'background';

const FOREGROUND_LANES = ['input', 'collab', 'interactive'] as const;
type ForegroundLane = (typeof FOREGROUND_LANES)[number];
const RANK: Record<SchedulerLane, number> = { input: 0, collab: 1, interactive: 2, background: 3 };

export const INPUT_SLICE_MS = 8;
export const IDLE_SLICE_MS = 24;
export const INPUT_ACTIVE_MS = 1000;

export interface SchedulerMessage {
  lane: ForegroundLane;
  run(): Promise<void> | void;
  userInput?: boolean;
  holdsIdleTasks?: boolean;
  reorderable?: boolean;
  mutates?: boolean;
  reframes?: boolean;
  key?: string;
  replaceableBy?: string;
  supersede?(): void;
}

export type TaskStep = 'done' | 'yield';

export interface SchedulerTask {
  kind: string;
  version: number;
  generation: number;
  run(budgetMs: number): Promise<TaskStep> | TaskStep;
  idleAfterInputMs?: number;
  onStale?(): 'cancel' | 'continue';
  cancel?(): void;
  fail?(error: unknown): void;
}

export interface PureTask<Input = unknown, Result = unknown> {
  kind: string;
  version: number;
  generation: number;
  input: Input;
  transfer?: Transferable[];
  compute(input: Input): Result;
  install(result: Result): Promise<void> | void;
  cancel?(reason?: unknown): void;
}

export interface TaskExecutor {
  run<Input, Result>(job: { kind: string; input: Input; transfer?: Transferable[] }): Promise<Result>;
}

export interface SchedulerHost {
  now(): number;
  turn(callback: () => void): void;
  timer(callback: () => void, ms: number): () => void;
  executor?: TaskExecutor;
  failed?(error: unknown): void;
}

interface QueuedMessage extends SchedulerMessage {
  seq: number;
}

export interface ResidentScheduler {
  submit(message: SchedulerMessage): void;
  schedule(task: SchedulerTask): void;
  dispatch<Input, Result>(task: PureTask<Input, Result>): void;
  budget(): number;
  readonly version: number;
  readonly generation: number;
  bump(changes: { version?: boolean; generation?: boolean }): void;
  pending(): { foreground: number; background: number };
}

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
