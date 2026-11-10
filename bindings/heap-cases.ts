import { expect, test } from 'bun:test';

/** The heap counters an editing engine exports, under their shared engine names. */
export interface HeapEngine {
  wasm_heap_counted(): boolean;
  wasm_live_bytes(): number;
  wasm_peak_bytes(): number;
  wasm_failed_allocation_bytes(): number;
  reset_wasm_peak_bytes(): void;
  set_wasm_heap_limit(bytes: number): void;
}

/** The engine's entry in `wasmModuleMemories()`. */
export interface HeapMemoryEntry {
  bufferBytes: number;
  maximumBytes?: number;
  liveBytes?: number;
  peakBytes?: number;
  failedAllocationBytes?: number;
  native?: true;
}

/**
 * What a limit does on each engine: wasm refuses the allocation past it; native
 * does not enforce it yet, so the allocation succeeds and nothing is recorded.
 */
export const HEAP_LIMIT = { wasm: 'refuses', native: 'not enforced' } as const;

function figures(engine: HeapEngine) {
  return {
    live: engine.wasm_live_bytes(),
    peak: engine.wasm_peak_bytes(),
    failed: engine.wasm_failed_allocation_bytes(),
  };
}

/**
 * Registers the shared heap reporting table for one engine. Byte counts differ
 * between the allocators, so the table compares behaviour, not sizes.
 * `hold` makes engine state that keeps heap until `release`; `entry` reads the
 * engine's memory listing.
 */
export function describeHeapTable<T>(
  engineKind: 'wasm' | 'native',
  engine: () => HeapEngine,
  hold: () => T,
  release: (held: T) => void,
  entry: () => HeapMemoryEntry | undefined,
) {
  test('the heap is counted', () => {
    expect(engine().wasm_heap_counted()).toBe(true);
  });

  test('every figure is a non-negative safe integer and nothing failed without a limit', () => {
    const { live, peak, failed } = figures(engine());
    for (const value of [live, peak, failed]) {
      expect(Number.isSafeInteger(value)).toBe(true);
      expect(value).toBeGreaterThanOrEqual(0);
    }
    expect(peak).toBeGreaterThanOrEqual(live);
    expect(failed).toBe(0);
  });

  test('live grows while engine state is held and drops when it is released', () => {
    const before = figures(engine());
    const held = hold();
    const holding = figures(engine());
    release(held);
    const after = figures(engine());
    expect(holding.live).toBeGreaterThan(before.live);
    expect(after.live).toBeLessThan(holding.live);
    expect(holding.peak).toBeGreaterThanOrEqual(holding.live);
    expect(after.peak).toBeGreaterThanOrEqual(holding.peak);
  });

  test('reset starts the peak window at the live bytes', () => {
    release(hold());
    const raised = figures(engine());
    expect(raised.peak).toBeGreaterThan(raised.live);
    engine().reset_wasm_peak_bytes();
    const reset = figures(engine());
    expect(reset.peak).toBe(reset.live);
    expect(reset.peak).toBeLessThan(raised.peak);
  });

  test('the memory listing carries the counters', () => {
    const listed = entry();
    expect(listed).toBeDefined();
    const { liveBytes, peakBytes, failedAllocationBytes, bufferBytes, maximumBytes } = listed!;
    for (const value of [liveBytes, peakBytes, failedAllocationBytes, bufferBytes]) {
      expect(Number.isSafeInteger(value)).toBe(true);
    }
    expect(peakBytes!).toBeGreaterThanOrEqual(liveBytes!);
    if (engineKind === 'native') {
      expect(bufferBytes).toBe(peakBytes!);
      expect(maximumBytes).toBeUndefined();
      expect(listed!.native).toBe(true);
    } else {
      expect(bufferBytes).toBeGreaterThanOrEqual(liveBytes!);
      expect(listed!.native).toBeUndefined();
    }
  });

  test(`a heap limit on ${engineKind}: ${HEAP_LIMIT[engineKind]}`, () => {
    const live = engine().wasm_live_bytes();
    engine().set_wasm_heap_limit(live + 1024);
    let thrown: unknown;
    try {
      release(hold());
    } catch (error) {
      thrown = error;
    } finally {
      engine().set_wasm_heap_limit(Infinity);
    }
    if (HEAP_LIMIT[engineKind] === 'refuses') {
      expect(thrown).toBeDefined();
      expect(engine().wasm_failed_allocation_bytes()).toBeGreaterThan(0);
    } else {
      expect(thrown).toBeUndefined();
      expect(engine().wasm_failed_allocation_bytes()).toBe(0);
    }
  });
}
