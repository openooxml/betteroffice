import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { WasmModuleMemory } from '@betteroffice/docx/yrs';
import {
  memoryPressureLevel,
  useMemoryPressure,
  type DocxMemoryPressure,
  type DocxMemoryStats,
} from './memoryStats';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook } = await import('@testing-library/react');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const GIB = 1024 ** 3;
const stats = (worker: WasmModuleMemory[] | null, main: WasmModuleMemory[] = []): DocxMemoryStats => ({
  limitBytes: 4 * GIB,
  main,
  worker,
});
const edit = (liveBytes: number, bufferBytes = liveBytes): WasmModuleMemory => ({
  label: 'docx-edit',
  bufferBytes,
  liveBytes,
  peakBytes: liveBytes,
  failedAllocationBytes: 0,
});

test('the level follows the fullest memory against 75% and 90% of the limit', () => {
  expect(memoryPressureLevel(stats([edit(2.9 * GIB)]))).toBe('normal');
  expect(memoryPressureLevel(stats([edit(3.1 * GIB)]))).toBe('warning');
  expect(memoryPressureLevel(stats(null, [edit(3.7 * GIB)]))).toBe('critical');
  expect(memoryPressureLevel(stats([edit(1 * GIB)], [edit(3.7 * GIB)]))).toBe('critical');
});

test('allocated bytes count, not the grown buffer, when the allocator reports them', () => {
  expect(memoryPressureLevel(stats([edit(1 * GIB, 3.9 * GIB)]))).toBe('normal');
  expect(memoryPressureLevel(stats([{ label: 'docx-layout', bufferBytes: 3.9 * GIB }]))).toBe(
    'critical'
  );
});

test('a budget moves the levels', () => {
  const budget = { warningBytes: 1 * GIB, criticalBytes: 2 * GIB };
  expect(memoryPressureLevel(stats([edit(1.5 * GIB)]), budget)).toBe('warning');
  expect(memoryPressureLevel(stats([edit(2.5 * GIB)]), budget)).toBe('critical');
  expect(memoryPressureLevel(stats([edit(0.5 * GIB)]), { criticalBytes: 0.25 * GIB })).toBe(
    'critical'
  );
});

test('pressure is reported on level changes only, and never while memory stays low', () => {
  const reports: DocxMemoryPressure[] = [];
  let worker: WasmModuleMemory[] = [edit(1 * GIB)];
  const workerMemory = () => worker;
  const hook = renderHook(
    ({ tick }: { tick: number }) =>
      useMemoryPressure((pressure) => reports.push(pressure), undefined, workerMemory, [tick]),
    { initialProps: { tick: 0 } }
  );
  hook.rerender({ tick: 1 });
  expect(reports).toEqual([]);

  worker = [edit(3.2 * GIB)];
  hook.rerender({ tick: 2 });
  hook.rerender({ tick: 3 });
  worker = [edit(3.8 * GIB)];
  hook.rerender({ tick: 4 });
  worker = [edit(1 * GIB)];
  hook.rerender({ tick: 5 });

  expect(reports.map((report) => report.level)).toEqual(['warning', 'critical', 'normal']);
  expect(reports[1].stats.worker).toEqual([edit(3.8 * GIB)]);
  expect(reports[1].stats.limitBytes).toBe(4 * GIB);
});

test('a callback supplied after mount hears the current level', () => {
  const reports: DocxMemoryPressure[] = [];
  const workerMemory = () => [edit(3.8 * GIB)];
  const hook = renderHook(
    ({ report }: { report?: (pressure: DocxMemoryPressure) => void }) =>
      useMemoryPressure(report, undefined, workerMemory, [0]),
    { initialProps: {} as { report?: (pressure: DocxMemoryPressure) => void } }
  );
  hook.rerender({ report: (pressure) => reports.push(pressure) });
  expect(reports.map((report) => report.level)).toEqual(['critical']);
});

test('a worker that goes away without a new frame is heard', () => {
  const reports: DocxMemoryPressure[] = [];
  let worker: WasmModuleMemory[] | null = [edit(3.8 * GIB)];
  const workerMemory = () => worker;
  const report = (pressure: DocxMemoryPressure) => reports.push(pressure);
  const frame = {};
  const hook = renderHook(
    ({ failure }: { failure: Error | null }) =>
      useMemoryPressure(report, undefined, workerMemory, [frame, failure]),
    { initialProps: { failure: null as Error | null } }
  );
  worker = null;
  hook.rerender({ failure: null });
  expect(reports.map((entry) => entry.level)).toEqual(['critical']);
  hook.rerender({ failure: new Error('out of memory') });
  expect(reports.map((entry) => entry.level)).toEqual(['critical', 'normal']);
});
