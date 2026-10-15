import { expect, test } from 'bun:test';
import { aggregate, modelRepetitions, newestFirst, repetitions, wilson, type Suite } from './stats';
import data from '../../../eval/results/v2.json';

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
