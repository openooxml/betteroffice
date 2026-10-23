import { expect, test } from 'bun:test';
import { aggregate, modelRepetitions, newestFirst, publishedSuite, repetitions, tiers, wilson, type Suite } from './stats';
import data from '../../../eval/results/v2.json';
import v3 from '../../../eval/results/v3.json';

const suite = data as Suite;
test('Wilson intervals at boundaries and a measured proportion', () => {
  expect(wilson(0, 10)[1]).toBeCloseTo(0.27753, 5);
  expect(wilson(10, 10)[0]).toBeCloseTo(0.72247, 5);
  expect(wilson(6, 10)[0]).toBeCloseTo(0.31267, 5);
  expect(wilson(6, 10)[1]).toBeCloseTo(0.83182, 5);
  expect(wilson(0, 0)).toEqual([0, 1]);
  expect(() => wilson(11, 10)).toThrow();
});

test('v2 aggregates are separated by model, arm and format', () => {
  expect(aggregate(suite, { model: 'gpt-6-luna', arm: 'generic', format: 'docx' }).passes).toBe(6);
  expect(aggregate(suite, { arm: 'generic' }).passes).toBe(14);
  expect(aggregate(suite, { arm: 'betteroffice' }).passes).toBe(30);
  expect(aggregate(suite, { model: 'absent' }).rate).toBeNull();
});

test('tier filters count runs and treat absent tiers as basic', () => {
  const tasks = suite.tasks.slice(0, 3).map((task, index) => ({ ...task, tier: index === 0 ? undefined : index === 1 ? 'basic' as const : 'complex' as const }));
  const runs = tasks.flatMap(task => [1, 2, 3].map(run => ({ ...suite.runs[0]!, task: task.id, model: 'model-a', arm: 'generic', run, pass: task.tier !== 'complex' || run === 1 })));
  const future = { ...suite, tasks, runs };
  expect(tiers(future)).toEqual(['basic', 'complex']);
  const basic = aggregate(future, { model: 'model-a', arm: 'generic', format: tasks[0]!.format, tier: 'basic' });
  const complex = aggregate(future, { model: 'model-a', arm: 'generic', tier: 'complex' });
  expect([basic.passes, basic.n, basic.tasks]).toEqual([6, 6, 2]);
  expect([complex.passes, complex.n, complex.tasks]).toEqual([1, 3, 1]);
  expect(aggregate(future).n).toBe(9);
  expect(aggregate(suite, { tier: 'basic' })).toEqual(aggregate(suite));
  expect(aggregate(suite, { tier: 'complex' }).n).toBe(0);
});

test('repeated runs and a second model subset keep their own denominators', () => {
  const tasks = suite.tasks.slice(0, 2);
  const runs = tasks.flatMap(t => [1, 2, 3].map(run => ({ ...suite.runs[0]!, task: t.id, arm: 'third', model: 'model-a', run, pass: !(t.id === tasks[0]!.id && run === 2) })));
  runs.push({ ...runs[0]!, model: 'model-b', pass: false });
  const future = { ...suite, suite: 'v3', tasks, runs };
  const a = aggregate(future, { model: 'model-a', arm: 'third' });
  expect([a.passes, a.n, a.tasks, a.allPass, a.allPassRate]).toEqual([5, 6, 2, 1, 0.5]);
  expect(a.runsPerTask).toEqual([3]);
  expect(aggregate(future, { model: 'model-b' }).n).toBe(1);
  expect(repetitions([1])).toBe('one run per task');
  expect(repetitions([1, 3])).toBe('1 to 3 runs per task');
  expect(newestFirst([suite, { ...future, date: '2026-10-16' }])[0]!.suite).toBe('v3');
});

test('per-model repetitions do not combine model denominators', () => {
  const tasks = suite.tasks.slice(0, 2);
  const runs = tasks.flatMap(t => [1, 2, 3].map(run => ({ ...suite.runs[0]!, task: t.id, model: 'model-a', run })));
  runs.push({ ...runs[0]!, model: 'model-b' });
  expect(modelRepetitions({ ...suite, tasks, runs })).toBe('3 runs per task with model-a; one run per task with model-b on 1 task');
  expect(modelRepetitions(suite)).toBe('one run per task');
});

test('numeric suite version precedes date, with date breaking version ties', () => {
  const v3 = { ...suite, suite: 'v3', date: '2026-01-01' };
  const v10 = { ...suite, suite: 'v10', date: '2025-01-01' };
  const older = { ...v3, date: '2025-01-01' };
  expect(newestFirst([suite, older, v10, v3]).map(s => [s.suite, s.date])).toEqual([
    ['v10', '2025-01-01'], ['v3', '2026-01-01'], ['v3', '2025-01-01'], ['v2', suite.date],
  ]);
});

test('publication allowlist preserves source data and filters tasks, runs and metadata', () => {
  const source = structuredClone(v3) as Suite;
  const before = structuredClone(source);
  const basic = publishedSuite(source);
  expect(tiers(basic)).toEqual(['basic']);
  expect(basic.tasks).toHaveLength(96);
  expect(basic.runs).toHaveLength(576);
  expect(aggregate(basic, { arm: 'betteroffice' }).passes).toBe(228);
  expect(aggregate(basic, { arm: 'generic' }).passes).toBe(161);
  expect(basic.fixtures.count).toBe(90);
  expect(basic.fixtures.description.join(' ')).not.toMatch(/complex/i);
  expect(basic.reproduce).toHaveLength(2);
  expect(basic.checker_validation).toBeUndefined();
  expect(publishedSuite(source, ['basic', 'complex'])).toEqual(source);
  expect(tiers(publishedSuite(source, ['complex']))).toEqual(['complex']);
  expect(publishedSuite(source, []).runs).toEqual([]);
  expect(source).toEqual(before);
});
