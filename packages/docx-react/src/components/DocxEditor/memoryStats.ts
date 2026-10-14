import { useEffect, useRef } from 'react';
import {
  WASM32_MEMORY_LIMIT_BYTES,
  wasmModuleMemories,
  type WasmModuleMemory,
} from '@betteroffice/docx/yrs';

/** The editor's wasm memories on the main thread and in its resident worker. */
export interface DocxMemoryStats {
  /** The most one wasm32 memory can hold: 4 GiB. */
  limitBytes: number;
  main: WasmModuleMemory[];
  /** The worker's memories as of its latest reply; null without a worker. */
  worker: WasmModuleMemory[] | null;
}

/**
 * How close the fullest wasm memory is to {@link DocxMemoryStats.limitBytes}: the
 * editing core's allocated bytes, and the memory size of modules that do not
 * count their heap.
 */
export type DocxMemoryPressureLevel = 'normal' | 'warning' | 'critical';

/** How much wasm memory the editor may use. */
export interface DocxMemoryBudget {
  /** Pressure is `warning` from here. Default: 75% of the 4 GiB limit. */
  warningBytes?: number;
  /** Pressure is `critical` from here. Default: 90% of the 4 GiB limit. */
  criticalBytes?: number;
  /**
   * The most the resident worker's editing core may allocate at once, applied
   * when a worker starts. An allocation past it fails as if the memory were
   * full. Default: no limit short of 4 GiB.
   */
  workerLimitBytes?: number;
}

export interface DocxMemoryPressure {
  level: DocxMemoryPressureLevel;
  stats: DocxMemoryStats;
}

export function readMemoryStats(
  workerMemory: () => WasmModuleMemory[] | null
): DocxMemoryStats {
  return {
    limitBytes: WASM32_MEMORY_LIMIT_BYTES,
    main: wasmModuleMemories(),
    worker: workerMemory(),
  };
}

/** Allocated bytes where the allocator reports them, the memory's size otherwise. */
function heldBytes(module: WasmModuleMemory): number {
  return module.liveBytes ?? module.bufferBytes;
}

export function memoryPressureLevel(
  stats: DocxMemoryStats,
  budget: DocxMemoryBudget = {}
): DocxMemoryPressureLevel {
  const critical = budget.criticalBytes ?? stats.limitBytes * 0.9;
  const warning = Math.min(budget.warningBytes ?? stats.limitBytes * 0.75, critical);
  let held = 0;
  for (const module of [...stats.main, ...(stats.worker ?? [])]) {
    held = Math.max(held, heldBytes(module));
  }
  if (held >= critical) return 'critical';
  if (held >= warning) return 'warning';
  return 'normal';
}

/**
 * Calls `onPressure` when the pressure level changes, checked whenever one of
 * `ticks` changes. Nothing is reported while the level stays `normal`.
 */
export function useMemoryPressure(
  onPressure: ((pressure: DocxMemoryPressure) => void) | undefined,
  budget: DocxMemoryBudget | undefined,
  workerMemory: () => WasmModuleMemory[] | null,
  ticks: readonly unknown[]
): void {
  const levelRef = useRef<DocxMemoryPressureLevel>('normal');
  const onPressureRef = useRef(onPressure);
  onPressureRef.current = onPressure;
  const warningBytes = budget?.warningBytes;
  const criticalBytes = budget?.criticalBytes;
  const reporting = onPressure !== undefined;
  useEffect(() => {
    const report = onPressureRef.current;
    if (!report) return;
    const stats = readMemoryStats(workerMemory);
    const level = memoryPressureLevel(stats, { warningBytes, criticalBytes });
    if (level === levelRef.current) return;
    levelRef.current = level;
    report({ level, stats });
  }, [workerMemory, warningBytes, criticalBytes, reporting, ...ticks]);
}
