import { publishedTiers } from '../../../shared/eval-publication';
import type { Run, Suite, Tier } from '../../../eval/formats/collect-results';
export type { Run, Suite, Tier } from '../../../eval/formats/collect-results';

export function wilson(passes: number, n: number): [number, number] {
  if (!Number.isInteger(n) || !Number.isInteger(passes) || n < 0 || passes < 0 || passes > n) throw new Error('Invalid pass counts');
  if (n === 0) return [0, 1];
  const z = 1.959963984540054, p = passes / n, denominator = 1 + z * z / n;
  const centre = (p + z * z / (2 * n)) / denominator;
  const radius = z * Math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / denominator;
  return [Math.max(0, centre - radius), Math.min(1, centre + radius)];
}

export function aggregate(suite: Suite, filter: { model?: string; arm?: string; format?: string; tier?: Tier } = {}) {
  const taskIds = new Set(suite.tasks.filter(t => (!filter.format || t.format === filter.format) && (!filter.tier || (t.tier ?? 'basic') === filter.tier)).map(t => t.id));
  const runs = suite.runs.filter(r => taskIds.has(r.task) && (!filter.model || r.model === filter.model) && (!filter.arm || r.arm === filter.arm));
  const groups = new Map<string, Run[]>();
  for (const r of runs) {
    const key = JSON.stringify([r.model, r.arm, r.task]);
    const group = groups.get(key) ?? [];
    group.push(r); groups.set(key, group);
  }
  const passes = runs.filter(r => r.pass).length;
  const tasks = groups.size;
  const allPass = [...groups.values()].filter(group => group.every(r => r.pass)).length;
  return { passes, n: runs.length, rate: runs.length ? passes / runs.length : null,
    interval: wilson(passes, runs.length), tasks, allPass,
    allPassRate: tasks ? allPass / tasks : null,
    runsPerTask: [...new Set([...groups.values()].map(group => group.length))].sort((a, b) => a - b) };
}

export const models = (suite: Suite) => [...new Set(suite.runs.map(r => r.model))].sort();
export const tiers = (suite: Suite): Tier[] => ['basic', 'complex'].filter(tier => suite.tasks.some(t => (t.tier ?? 'basic') === tier)) as Tier[];
export const formats = (suite: Suite) => [...new Set(suite.tasks.map(t => t.format))].sort((a, b) => {
  const order = ['docx', 'xlsx', 'pptx'];
  return (order.includes(a) ? order.indexOf(a) : order.length) - (order.includes(b) ? order.indexOf(b) : order.length) || a.localeCompare(b);
});
export const newestFirst = (suites: Suite[]) => [...suites].sort((a, b) => b.suite.localeCompare(a.suite, 'en', { numeric: true }) || b.date.localeCompare(a.date));
export const percent = (value: number) => `${(value * 100).toFixed(1)}%`;
export function repetitions(counts: number[]) {
  if (!counts.length) return 'No runs';
  if (counts.length === 1) return counts[0] === 1 ? 'one run per task' : `${counts[0]} runs per task`;
  return `${counts[0]} to ${counts.at(-1)} runs per task`;
}

export function modelRepetitions(suite: Suite) {
  const names = models(suite);
  return names.map(model => {
    const description = repetitions(aggregate(suite, { model }).runsPerTask);
    const taskCount = new Set(suite.runs.filter(r => r.model === model).map(r => r.task)).size;
    const subset = taskCount < suite.tasks.length ? ` on ${taskCount} ${taskCount === 1 ? 'task' : 'tasks'}` : '';
    return names.length === 1 && !subset ? description : `${description} with ${model}${subset}`;
  }).join('; ');
}

export function publishedSuite(suite: Suite, allowlist: readonly Tier[] = publishedTiers): Suite {
  const tasks = suite.tasks.filter(task => allowlist.includes(task.tier ?? 'basic'));
  const ids = new Set(tasks.map(task => task.id));
  const hidden = tiers(suite).filter(tier => !allowlist.includes(tier));
  const visibleCopy = (text: string) => hidden.every(tier => !new RegExp(`\\b${tier}\\b`, 'i').test(text));
  return {
    ...suite, tasks, runs: suite.runs.filter(run => ids.has(run.task)),
    fixtures: { ...suite.fixtures, count: tasks.length === suite.tasks.length ? suite.fixtures.count : new Set(tasks.map(task => task.fixture)).size, description: suite.fixtures.description.filter(visibleCopy) },
    reproduce: suite.reproduce.filter(command => allowlist.includes((command.match(/--tier\s+(basic|complex)\b/)?.[1] ?? 'basic') as Tier)),
    checker_validation: tasks.length === suite.tasks.length ? suite.checker_validation : undefined,
  };
}
