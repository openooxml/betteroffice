import { expect, test } from 'bun:test';
import { DEFAULT_PRIVATE_PARALLEL, UNBOUNDED, allocate, matrixLoad, parallelPlan, parallelSummary } from './parallel.mjs';

const sample = (id: string, format: string, pages: number, cells: number | null = null) => ({
  id, format, metadata: { source: { sha256: id }, reference: { pages }, ...(cells ? { calculation: { status: 'ok', result_cells: cells } } : {}) },
});
const corpus = () => ({
  formats: ['docx', 'pptx', 'xlsx'],
  samples: [
    ...Array.from({ length: 1776 }, (_, index) => sample(`d${index}`, 'docx', 1 + ((index * 37) % 22))),
    ...Array.from({ length: 656 }, (_, index) => sample(`p${index}`, 'pptx', 1 + ((index * 13) % 37))),
    ...Array.from({ length: 3297 }, (_, index) => sample(`x${index}`, 'xlsx', 1 + (index % 2), index % 5 ? 200 : null)),
  ],
});
const small = () => ({ formats: ['docx', 'pptx'], samples: [sample('a', 'docx', 3), sample('b', 'pptx', 4)] });
type Load = Record<string, { jobs: number; minutes: number }>;
const runningOf = (load: Load) => Object.keys(load).filter((name) => load[name].jobs > 0 && load[name].minutes > 0);
const wallOf = (load: Load, share: Record<string, number>) => Math.max(...runningOf(load).map((name) => load[name].minutes / share[name]));

function random(seed: number) {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = Math.imul(state ^ (state >>> 15), state | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

function expectInvariants(load: Load, total: number, share: Record<string, number>) {
  const running = runningOf(load);
  const used = running.reduce((sum, name) => sum + share[name], 0);
  expect(used).toBeLessThanOrEqual(total);
  for (const name of Object.keys(share)) expect(share[name]).toBeGreaterThanOrEqual(1);
  for (const name of running) expect(share[name]).toBeLessThanOrEqual(load[name].jobs);
  if (used < total) for (const name of running) expect(share[name]).toBe(load[name].jobs);
}

test('the split never exceeds max_parallel, never starves or over-fills a matrix, and leaves no usable slot idle', () => {
  for (const plan of [corpus(), small()]) {
    const load = matrixLoad(plan);
    for (const total of [runningOf(load).length, DEFAULT_PRIVATE_PARALLEL, 13, 20, 50, 500])
      expectInvariants(load, total, parallelPlan(plan, { privateRun: true, total }));
  }
});

function optimum(load: Load, total: number) {
  const running = runningOf(load);
  const candidates = running.flatMap((name) => Array.from({ length: load[name].jobs }, (_, index) => load[name].minutes / (index + 1)));
  const feasible = (limit: number) => {
    let used = 0;
    for (const name of running) {
      const slots = Math.max(1, Math.ceil(load[name].minutes / limit - 1e-9));
      if (slots > load[name].jobs) return false;
      used += slots;
    }
    return used <= total;
  };
  return Math.min(...candidates.filter(feasible));
}

test('the split reaches the exact optimum for the longest matrix, including where single slot moves cannot find it', () => {
  const load: Load = { a: { jobs: 2, minutes: 10 }, b: { jobs: 2, minutes: 10 }, c: { jobs: 3, minutes: 1 } };
  expect(wallOf(load, allocate(load, 5))).toBe(5);
  const next = random(7);
  for (let round = 0; round < 300; round += 1) {
    const sized: Load = Object.fromEntries(['a', 'b', 'c', 'd'].map((name) => {
      const jobs = 1 + Math.floor(next() * 6);
      return [name, { jobs, minutes: 1 + Math.floor(next() * 50) }];
    }));
    const total = 4 + Math.floor(next() * 20);
    expect(wallOf(sized, allocate(sized, total))).toBeCloseTo(optimum(sized, total), 9);
  }
});

test('property: for any loads and job counts the invariants hold and no single slot move shortens the longest matrix', () => {
  const next = random(20261110);
  const names = ['measure', 'docx-benchmark', 'xlsx-benchmark', 'pptx-benchmark', 'xlsx-fidelity', 'roundtrip-a', 'roundtrip-b', 'timing'];
  for (let round = 0; round < 500; round += 1) {
    const load: Load = Object.fromEntries(names.map((name) => {
      const jobs = next() < 0.15 ? 0 : 1 + Math.floor(next() * 300);
      return [name, { jobs, minutes: jobs ? 1 + next() * 3000 : 0 }];
    }));
    const running = runningOf(load);
    if (!running.length) continue;
    const total = running.length + Math.floor(next() * 40);
    const share = allocate(load, total);
    expectInvariants(load, total, share);
    for (const from of running.filter((name) => share[name] > 1))
      for (const to of running.filter((name) => name !== from && share[name] < load[name].jobs))
        expect(wallOf(load, { ...share, [from]: share[from] - 1, [to]: share[to] + 1 })).toBeGreaterThanOrEqual(wallOf(load, share) - 1e-9);
  }
});

test('a larger cap never projects a longer run, and the largest matrix gets more slots than the smallest', () => {
  const plan = corpus();
  const load = matrixLoad(plan);
  let previous = Infinity;
  for (let total = runningOf(load).length; total <= 40; total += 1) {
    const wall = wallOf(load, parallelPlan(plan, { privateRun: true, total }));
    expect(wall).toBeLessThanOrEqual(previous + 1e-9);
    previous = wall;
  }
  const share = parallelPlan(plan, { privateRun: true, total: DEFAULT_PRIVATE_PARALLEL });
  expect(share.measure).toBeGreaterThan(share.timing);
});

test('public runs keep unbounded parallelism, and a cap below the matrix count fails in planning', () => {
  expect(Object.values(parallelPlan(corpus(), { privateRun: false })).every((value) => value === UNBOUNDED)).toBe(true);
  const running = runningOf(matrixLoad(corpus())).length;
  expect(() => parallelPlan(corpus(), { privateRun: true, total: running - 1 })).toThrow('below the');
  expect(() => parallelPlan(corpus(), { privateRun: true, total: 0 })).toThrow('positive whole number');
  expect(() => parallelPlan(corpus(), { privateRun: true, total: 2.5 })).toThrow('positive whole number');
});

test('the printed split shows each matrix wall time and the projected wall of the run, with numbers only', () => {
  const plan = corpus();
  const lines = parallelSummary(plan, parallelPlan(plan, { privateRun: true, total: 12 }));
  expect(lines.length).toBeGreaterThan(4);
  expect(lines.join('\n')).not.toMatch(/\b[dpx]\d+\b/);
  for (const line of lines.slice(0, -1))
    expect(line).toMatch(/^[a-z-]+: at most \d+ at once for \d+ jobs?, ~\d+ runner-minutes planned, about \d+\.\d h wall$/);
  expect(lines.at(-1)).toMatch(/^projected wall of the benchmark matrices: about \d+\.\d h for \d+ jobs$/);
});
