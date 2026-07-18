/**
 * Opt-in per-keystroke instrumentation for the JS/WASM editor hot path.
 *
 * The seam is deliberately inert unless the page was opened with
 * `?perfTrace=1`. Samples are kept in memory for the Playwright profiler and
 * exposed through `window.__perfTrace`; production builds pay one null check
 * at an instrumented boundary and allocate nothing.
 */

export const PERF_TRACE_STAGES = [
  'inputQueued',
  'inputAck',
  'yrsOp',
  'storyBlocks',
  'measure',
  'paginate',
  'displayList',
  'engineWorker',
  'engineWorkerTotal',
  'engineSelection',
  'engineEdit',
  'engineLower',
  'engineMeasure',
  'enginePaginate',
  'engineDisplayInput',
  'engineDisplayBuild',
  'engineDisplayFinalize',
  'engineDisplay',
  'engineEncode',
  'workerRoundtrip',
  'deltaDecode',
  'canvasReplay',
  'selectionReads',
  'selectionContext',
] as const;

export type PerfTraceStage = (typeof PERF_TRACE_STAGES)[number];

export interface PerfTraceMetadata {
  bytes?: number;
  inputBytes?: number;
  outputBytes?: number;
  calls?: number;
  detail?: string;
}

export interface PerfTraceSample extends PerfTraceMetadata {
  stage: PerfTraceStage;
  durationMs: number;
  /** `performance.now()` at sample completion. */
  atMs: number;
  /** Monotonic input sequence; zero means startup/non-keystroke work. */
  keystroke: number;
}

export interface PerfTraceAggregate {
  count: number;
  totalMs: number;
  p50Ms: number;
  p95Ms: number;
  maxMs: number;
  totalBytes: number;
  totalInputBytes: number;
  totalOutputBytes: number;
}

export interface PerfTraceSnapshot {
  enabled: true;
  currentKeystroke: number;
  frameDeltas: { fullRecovery: number; delta: number };
  samples: PerfTraceSample[];
  aggregates: Partial<Record<PerfTraceStage, PerfTraceAggregate>>;
}

export interface PerfTraceController {
  readonly enabled: true;
  readonly samples: PerfTraceSample[];
  readonly aggregates: Partial<Record<PerfTraceStage, PerfTraceAggregate>>;
  readonly frameDeltas: { fullRecovery: number; delta: number };
  currentKeystroke: number;
  beginKeystroke(detail?: string): number;
  record(stage: PerfTraceStage, durationMs: number, metadata?: PerfTraceMetadata): void;
  recordForKeystroke(
    keystroke: number,
    stage: PerfTraceStage,
    durationMs: number,
    metadata?: PerfTraceMetadata
  ): void;
  recordFrameDelta(fullRecovery: boolean): void;
  reset(): void;
  snapshot(): PerfTraceSnapshot;
}

declare global {
  // eslint-disable-next-line no-var
  var __perfTrace: PerfTraceController | undefined;
}

function percentile(values: number[], p: number): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const index = Math.min(sorted.length - 1, Math.ceil((p / 100) * sorted.length) - 1);
  return sorted[index] ?? 0;
}

function refreshAggregate(controller: PerfTraceController, stage: PerfTraceStage): void {
  const samples = controller.samples.filter((sample) => sample.stage === stage);
  const durations = samples.map((sample) => sample.durationMs);
  controller.aggregates[stage] = {
    count: samples.length,
    totalMs: durations.reduce((sum, duration) => sum + duration, 0),
    p50Ms: percentile(durations, 50),
    p95Ms: percentile(durations, 95),
    maxMs: Math.max(0, ...durations),
    totalBytes: samples.reduce((sum, sample) => sum + (sample.bytes ?? 0), 0),
    totalInputBytes: samples.reduce((sum, sample) => sum + (sample.inputBytes ?? 0), 0),
    totalOutputBytes: samples.reduce((sum, sample) => sum + (sample.outputBytes ?? 0), 0),
  };
}

function traceRequested(): boolean {
  return (
    typeof window !== 'undefined' &&
    new URLSearchParams(window.location.search).get('perfTrace') === '1'
  );
}

export function getPerfTrace(): PerfTraceController | undefined {
  if (!traceRequested()) return undefined;
  if (globalThis.__perfTrace) return globalThis.__perfTrace;

  const controller: PerfTraceController = {
    enabled: true,
    samples: [],
    aggregates: {},
    frameDeltas: { fullRecovery: 0, delta: 0 },
    currentKeystroke: 0,
    beginKeystroke(detail): number {
      controller.currentKeystroke += 1;
      controller.recordForKeystroke(controller.currentKeystroke, 'inputQueued', 0, {
        calls: 1,
        ...(detail ? { detail } : {}),
      });
      return controller.currentKeystroke;
    },
    record(stage, durationMs, metadata = {}): void {
      controller.recordForKeystroke(controller.currentKeystroke, stage, durationMs, metadata);
    },
    recordForKeystroke(keystroke, stage, durationMs, metadata = {}): void {
      if (!Number.isFinite(durationMs) || durationMs < 0) return;
      controller.samples.push({
        stage,
        durationMs,
        atMs: performance.now(),
        keystroke,
        ...metadata,
      });
      refreshAggregate(controller, stage);
    },
    recordFrameDelta(fullRecovery): void {
      if (fullRecovery) controller.frameDeltas.fullRecovery += 1;
      else controller.frameDeltas.delta += 1;
    },
    reset(): void {
      controller.samples.length = 0;
      for (const stage of PERF_TRACE_STAGES) delete controller.aggregates[stage];
      controller.currentKeystroke = 0;
      controller.frameDeltas.fullRecovery = 0;
      controller.frameDeltas.delta = 0;
    },
    snapshot(): PerfTraceSnapshot {
      return {
        enabled: true,
        currentKeystroke: controller.currentKeystroke,
        frameDeltas: { ...controller.frameDeltas },
        samples: controller.samples.map((sample) => ({ ...sample })),
        aggregates: Object.fromEntries(
          Object.entries(controller.aggregates).map(([stage, aggregate]) => [
            stage,
            aggregate ? { ...aggregate } : aggregate,
          ])
        ),
      };
    },
  };
  globalThis.__perfTrace = controller;
  return controller;
}

export function beginPerfKeystroke(detail?: string): number {
  return getPerfTrace()?.beginKeystroke(detail) ?? 0;
}

export function tracePerfSync<T>(
  stage: PerfTraceStage,
  operation: () => T,
  metadata?: PerfTraceMetadata | ((result: T) => PerfTraceMetadata)
): T {
  const trace = getPerfTrace();
  if (!trace) return operation();
  const started = performance.now();
  const result = operation();
  trace.record(
    stage,
    performance.now() - started,
    typeof metadata === 'function' ? metadata(result) : metadata
  );
  return result;
}

export async function tracePerfAsync<T>(
  stage: PerfTraceStage,
  operation: () => Promise<T>,
  metadata?: PerfTraceMetadata | ((result: T) => PerfTraceMetadata),
  keystroke?: number
): Promise<T> {
  const trace = getPerfTrace();
  if (!trace) return operation();
  const started = performance.now();
  const result = await operation();
  const resolvedMetadata = typeof metadata === 'function' ? metadata(result) : metadata;
  if (keystroke !== undefined) {
    trace.recordForKeystroke(keystroke, stage, performance.now() - started, resolvedMetadata);
  } else {
    trace.record(stage, performance.now() - started, resolvedMetadata);
  }
  return result;
}

// Initialize the public seam as soon as an editor instrumentation site loads.
getPerfTrace();
