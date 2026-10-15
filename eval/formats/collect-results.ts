import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';

export type Task = { id: string; format: string; title: string; prompt: string };
export type Arm = { id: string; label: string; summary: string; tool_calls: string };
export type Run = {
  model: string; arm: string; task: string; run: number; date: string;
  pass: boolean; check: string; reason: string | null; cause: 'library' | 'model' | null;
  tool_calls: number; seconds: number;
};
export type Suite = {
  schema: 1; suite: string; date: string;
  fixtures: { kind: string; count: number; url: string; metadata: string };
  harness: string; arms: Arm[]; tasks: Task[]; runs: Run[];
  checker_validation?: { accepted: number; untouched_rejected: number; corrupted_rejected: number };
};
export type Reason = Pick<Run, 'model' | 'arm' | 'task' | 'run' | 'reason' | 'cause'>;
export type Report = { model: string; results: {
  task: string; pass: boolean; check_message: string; tool_calls: number; duration_seconds: number;
}[] };
export type Group = { arm: string; directory: string; run: number; date: string };
export const defaultArms: Arm[] = [
  { id: 'generic', label: 'Generic file tools', summary: 'Shell and Python with python-docx, openpyxl and python-pptx, without network access.', tool_calls: 'Shell commands and file edits; one command can run a whole script.' },
  { id: 'betteroffice', label: 'BetterOffice MCP', summary: 'Only the BetterOffice MCP server for document access.', tool_calls: 'BetterOffice MCP calls.' },
];
const order = (a: string, b: string) => a < b ? -1 : a > b ? 1 : 0;
const key = (r: Pick<Run, 'model' | 'arm' | 'task' | 'run'>) => JSON.stringify([r.model, r.arm, r.task, r.run]);
const validDate = (date: string) => /^\d{4}-\d{2}-\d{2}$/.test(date) && new Date(date).toISOString().slice(0, 10) === date;

export function parseArgs(args: string[]) {
  const groups: Group[] = [];
  let suite = '', reasons = '', output = '';
  let arms = '', fixtures = resolve(import.meta.dir, 'fixtures.json');
  for (let i = 0; i < args.length; i += 2) {
    const flag = args[i], value = args[i + 1];
    if (!value || value.startsWith('--')) throw new Error(`Missing value for ${flag}`);
    if (flag === '--runs') {
      const split = value.indexOf('=');
      if (split < 1 || split === value.length - 1) throw new Error('Expected --runs <arm>=<directory>');
      groups.push({ arm: value.slice(0, split), directory: value.slice(split + 1), run: 1, date: '' });
    } else if (flag === '--run-index' || flag === '--date') {
      const group = groups.at(-1);
      if (!group) throw new Error(`${flag} must follow --runs`);
      if (flag === '--date') group.date = value;
      else group.run = Number(value);
    } else if (flag === '--suite') suite = value;
    else if (flag === '--reasons') reasons = value;
    else if (flag === '--output') output = value;
    else if (flag === '--arms') arms = value;
    else if (flag === '--fixtures') fixtures = value;
    else throw new Error(`Unknown option ${flag}`);
  }
  if (!suite || !groups.length) throw new Error('Required: --suite and --runs with --date per group');
  for (const g of groups) {
    if (!validDate(g.date) || !Number.isInteger(g.run) || g.run < 1) throw new Error('Invalid group date or run index');
  }
  return { groups, suite, reasons, output, arms, fixtures };
}

export function collectResults(options: {
  suite: string; tasks: Task[]; arms: Arm[]; fixtures: Suite['fixtures']; reasons: Reason[];
  groups: (Group & { report: Report })[]; checker_validation?: Suite['checker_validation'];
}): Suite {
  const tasks = new Map(options.tasks.map(t => [t.id, t]));
  const arms = new Map(options.arms.map(a => [a.id, a]));
  if (tasks.size !== options.tasks.length || arms.size !== options.arms.length) throw new Error('Duplicate task or arm definition');
  const reasons = new Map<string, Reason>();
  for (const r of options.reasons) {
    if (!tasks.has(r.task) || !arms.has(r.arm)) throw new Error(`Unknown reason task or arm: ${r.task}/${r.arm}`);
    if (reasons.has(key(r))) throw new Error(`Duplicate reason: ${key(r)}`);
    if (!r.reason?.trim() || !['library', 'model'].includes(r.cause ?? '')) throw new Error(`Invalid failure reason: ${r.task}`);
    reasons.set(key(r), r);
  }
  const runs: Run[] = [], seen = new Set<string>(), used = new Set<string>();
  for (const g of options.groups) {
    if (!arms.has(g.arm)) throw new Error(`Unknown arm: ${g.arm}; supply --arms for additional arms`);
    if (!validDate(g.date) || !Number.isInteger(g.run) || g.run < 1 || !g.report.model) throw new Error('Invalid run group');
    for (const r of g.report.results) {
      if (!tasks.has(r.task)) throw new Error(`Unknown task: ${r.task}`);
      if (typeof r.pass !== 'boolean' || typeof r.check_message !== 'string' || !Number.isInteger(r.tool_calls) || r.tool_calls < 0 || !Number.isFinite(r.duration_seconds) || r.duration_seconds < 0) throw new Error(`Invalid result: ${r.task}`);
      const identity = { model: g.report.model, arm: g.arm, task: r.task, run: g.run };
      const id = key(identity), reason = reasons.get(id);
      if (seen.has(id)) throw new Error(`Duplicate run: ${id}`);
      seen.add(id); used.add(r.task);
      if (!r.pass && options.suite === 'v2' && !reason) throw new Error(`Failed v2 run without a reason: ${id}`);
      runs.push({ ...identity, date: g.date, pass: r.pass, check: r.check_message,
        reason: r.pass ? null : reason?.reason ?? null, cause: r.pass ? null : reason?.cause ?? null,
        tool_calls: r.tool_calls, seconds: r.duration_seconds });
    }
  }
  if (!runs.length) throw new Error('No results');
  runs.sort((a, b) => order(a.model, b.model) || options.arms.findIndex(x => x.id === a.arm) - options.arms.findIndex(x => x.id === b.arm) || order(a.task, b.task) || a.run - b.run);
  return { schema: 1, suite: options.suite, date: runs.map(r => r.date).sort().at(-1)!,
    fixtures: options.fixtures, harness: 'codex exec',
    arms: options.arms.filter(a => runs.some(r => r.arm === a.id)),
    tasks: [...used].sort(order).map(id => {
      const t = tasks.get(id)!;
      return { id: t.id, format: t.format, title: t.title, prompt: t.prompt };
    }), runs, ...(options.checker_validation ? { checker_validation: options.checker_validation } : {}) };
}

export const serialize = (suite: Suite) => JSON.stringify(suite, null, 2) + '\n';

export async function main(args: string[]) {
  const options = parseArgs(args);
  const tasks: (Task & { output: string })[] = JSON.parse(await readFile(resolve(import.meta.dir, 'tasks.json'), 'utf8'));
  const manifest = JSON.parse(await readFile(options.fixtures, 'utf8'));
  const arms: Arm[] = options.arms ? JSON.parse(await readFile(options.arms, 'utf8')) : defaultArms;
  const reasons: Reason[] = options.reasons ? JSON.parse(await readFile(options.reasons, 'utf8')) : [];
  const groups = [];
  for (const g of options.groups) {
    const report: Omit<Report, 'results'> & { results: (Report['results'][number] & { agent_exit_code?: number; timed_out?: boolean })[] } = JSON.parse(await readFile(resolve(g.directory, 'results.json'), 'utf8'));
    for (const r of report.results) {
      const task = tasks.find(t => t.id === r.task);
      if (!task) throw new Error(`Unknown task: ${r.task}`);
      const check = Bun.spawnSync(['python3', resolve(import.meta.dir, 'check.py'), r.task, resolve(g.directory, r.task, task.output)], { env: process.env });
      const message = check.stdout.toString().trim();
      const prefix = `${check.exitCode === 0 ? 'PASS' : 'FAIL'}: ${r.task}: `;
      if (![0, 1].includes(check.exitCode) || !message.startsWith(prefix)) throw new Error(`Checker could not score ${r.task}: ${check.stderr.toString()}`);
      if (r.timed_out || (r.agent_exit_code !== undefined && r.agent_exit_code !== 0) || r.tool_calls === 0) throw new Error(`Runner did not complete a valid workflow: ${r.task}`);
      r.pass = check.exitCode === 0;
      r.check_message = message.slice(prefix.length);
    }
    groups.push({ ...g, report });
  }
  const origin = manifest.origin.replace(/\/$/, '');
  const readme = await readFile(resolve(import.meta.dir, '../README.md'), 'utf8');
  const validation = readme.match(/\((\d+) accepted, (\d+) untouched rejected, (\d+)\s+corrupted rejected\)/);
  if (!validation) throw new Error('Missing checker validation counts in eval/README.md');
  const result = collectResults({ ...options, tasks, arms, reasons, groups,
    checker_validation: { accepted: Number(validation[1]), untouched_rejected: Number(validation[2]), corrupted_rejected: Number(validation[3]) },
    fixtures: { kind: manifest.kind ?? 'synthetic', count: Object.keys(manifest.fixtures).length, url: origin + '/', metadata: origin + '/metadata.json' } });
  if (options.output) await writeFile(options.output, serialize(result));
  else process.stdout.write(serialize(result));
}

if (import.meta.main) await main(process.argv.slice(2));
