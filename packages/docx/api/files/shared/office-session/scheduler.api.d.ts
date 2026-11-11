// @betteroffice/docx declaration shared/office-session/scheduler.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
export type SchedulerLane = 'input' | 'collab' | 'interactive' | 'background';
declare const FOREGROUND_LANES: readonly ["input", "collab", "interactive"];
type ForegroundLane = (typeof FOREGROUND_LANES)[number];
export declare const INPUT_SLICE_MS = 8;
export declare const IDLE_SLICE_MS = 24;
export declare const INPUT_ACTIVE_MS = 1000;
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
    run<Input, Result>(job: {
        kind: string;
        input: Input;
        transfer?: Transferable[];
    }): Promise<Result>;
}
export interface SchedulerHost {
    now(): number;
    turn(callback: () => void): void;
    timer(callback: () => void, ms: number): () => void;
    executor?: TaskExecutor;
    failed?(error: unknown): void;
}
export interface ResidentScheduler {
    submit(message: SchedulerMessage): void;
    schedule(task: SchedulerTask): void;
    dispatch<Input, Result>(task: PureTask<Input, Result>): void;
    budget(): number;
    readonly version: number;
    readonly generation: number;
    bump(changes: {
        version?: boolean;
        generation?: boolean;
    }): void;
    pending(): {
        foreground: number;
        background: number;
    };
}
export declare function createResidentScheduler(host: SchedulerHost): ResidentScheduler;
export {};
