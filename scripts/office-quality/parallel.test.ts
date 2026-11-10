import { expect, test } from 'bun:test';
import { DEFAULT_PRIVATE_PARALLEL, UNBOUNDED, matrixMinutes, parallelPlan, parallelSummary } from './parallel.mjs';

const sample = (id: string, format: string, pages: number, cells: number | null = null) => ({
  id, format, metadata: { source: { sha256: id }, reference: { pages }, ...(cells ? { calculation: { status: 'ok', result_cells: cells } } : {}) },
});
const corpus = (scale = 1) => ({
  formats: ['docx', 'pptx', 'xlsx'],
  samples: [
    ...Array.from({ length: 1776 * scale }, (_, index) => sample(`d${index}`, 'docx', 1 + ((index * 37) % 22))),
    ...Array.from({ length: 656 * scale }, (_, index) => sample(`p${index}`, 'pptx', 1 + ((index * 13) % 37))),
    ...Array.from({ length: 3297 * scale }, (_, index) => sample(`x${index}`, 'xlsx', 1 + (index % 2), index % 5 ? 200 : null)),
  ],
});
const small = () => ({ formats: ['docx', 'pptx'], samples: [sample('a', 'docx', 3), sample('b', 'pptx', 4)] });

test('a private run never holds more jobs than max_parallel and never gives a running matrix zero', () => {
  for (const plan of [corpus(), small()]) {
    const minutes = matrixMinutes(plan);
    const running = Object.keys(minutes).filter((name) => minutes[name] > 0);
    for (const total of [running.length, DEFAULT_PRIVATE_PARALLEL, 13, 20, 50]) {
      const share = parallelPlan(plan, { privateRun: true, total });
      expect(running.reduce((sum, name) => sum + share[name], 0)).toBeLessThanOrEqual(total);
      for (const name of Object.keys(share)) expect(share[name]).toBeGreaterThanOrEqual(1);
    }
  }
});

test('the split follows each matrix planned runner-minutes', () => {
  const share = parallelPlan(corpus(), { privateRun: true, total: DEFAULT_PRIVATE_PARALLEL });
  const minutes = matrixMinutes(corpus());
  expect(share.measure).toBeGreaterThan(share['xlsx-fidelity']);
  expect(share.measure).toBeGreaterThanOrEqual(share['pptx-benchmark']);
  expect(minutes.measure).toBeGreaterThan(minutes['xlsx-fidelity']);
});

test('public runs keep unbounded parallelism, and a cap below the matrix count fails in planning', () => {
  expect(Object.values(parallelPlan(corpus(), { privateRun: false })).every((value) => value === UNBOUNDED)).toBe(true);
  const running = Object.values(matrixMinutes(corpus())).filter((value) => value > 0).length;
  expect(() => parallelPlan(corpus(), { privateRun: true, total: running - 1 })).toThrow('below the');
  expect(() => parallelPlan(corpus(), { privateRun: true, total: 0 })).toThrow('positive whole number');
  expect(() => parallelPlan(corpus(), { privateRun: true, total: 2.5 })).toThrow('positive whole number');
});

test('the printed split names matrices and numbers only', () => {
  const plan = corpus();
  const lines = parallelSummary(plan, parallelPlan(plan, { privateRun: true, total: 12 }));
  expect(lines.length).toBeGreaterThan(4);
  expect(lines.join('\n')).not.toMatch(/\b[dpx]\d+\b/);
  for (const line of lines) expect(line).toMatch(/^[a-z-]+: at most \d+ at once, ~\d+ runner-minutes planned$/);
});
