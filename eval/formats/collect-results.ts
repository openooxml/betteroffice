import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';

export type Tier = 'basic' | 'complex';
export type Task = { id: string; tier?: Tier; format: string; title: string; prompt: string; fixture: string; bytes: number; sha256: string };
export type Arm = { id: string; label: string; comparison_label?: string; summary: string; details: string[]; tool_calls: string };
export type Run = {
  model: string; arm: string; task: string; run: number; date: string;
  pass: boolean; check: string; reason: string | null; cause: 'library' | 'model' | 'timeout' | 'agent-exit' | 'no-tool-calls' | null;
  tool_calls: number; seconds: number;
};
export type Suite = {
  schema: 1; suite: string; date: string;
  fixtures: { kind: string; count: number; url: string; metadata: string; description: string[] };
  harness: string; limits: string[]; reproduce: string[]; arms: Arm[]; tasks: Task[]; runs: Run[];
  checker_validation?: { accepted: number; untouched_rejected: number; corrupted_rejected: number };
};
export type Reason = Pick<Run, 'model' | 'arm' | 'task' | 'run' | 'reason' | 'cause'>;
export type Report = { model: string; run_directories?: boolean; results: {
  task: string; run?: number; pass: boolean; check_message: string; tool_calls: number; duration_seconds: number;
  agent_exit_code?: number | null; timed_out?: boolean;
}[] };
type SourceTask = Omit<Task, 'bytes' | 'sha256'> & { output: string; expected?: { answer?: unknown } };
export type Group = { arm: string; directory: string; run: number; date: string };
export type SuiteMetadata = Pick<Suite, 'harness' | 'limits' | 'reproduce' | 'arms' | 'checker_validation'> & {
  fixtures: Pick<Suite['fixtures'], 'kind' | 'description'>;
};
const requiredText = (value: unknown, field: string) => {
  if (typeof value !== 'string' || !value.trim()) throw new Error('Missing or invalid ' + field);
};
const requiredParagraphs = (value: unknown, field: string) => {
  if (!Array.isArray(value) || !value.length) throw new Error('Missing or invalid ' + field);
  for (const paragraph of value) requiredText(paragraph, field);
};

export function validateMetadata(value: SuiteMetadata) {
  requiredText(value.harness, 'harness');
  requiredParagraphs(value.limits, 'limits');
  requiredParagraphs(value.reproduce, 'reproduce');
  requiredText(value.fixtures?.kind, 'fixtures.kind');
  requiredParagraphs(value.fixtures?.description, 'fixtures.description');
  if (!Array.isArray(value.arms) || !value.arms.length) throw new Error('Missing arms');
  for (const arm of value.arms) {
    for (const field of ['id', 'label', 'summary', 'tool_calls'] as const) requiredText(arm?.[field], 'arms.' + field);
    requiredParagraphs(arm?.details, 'arms.details');
    if (arm.comparison_label !== undefined) requiredText(arm.comparison_label, 'arms.comparison_label');
  }
  if (value.checker_validation !== undefined) {
    for (const field of ['accepted', 'untouched_rejected', 'corrupted_rejected'] as const) {
      const count = value.checker_validation?.[field];
      if (!Number.isInteger(count) || count < 0) throw new Error('Invalid checker_validation.' + field);
    }
  }
}

export function validateSuite(value: Suite) {
  if (!value || value.schema !== 1 || !/^v\d+$/.test(value.suite) || !validDate(value.date)) throw new Error('Invalid suite');
  validateMetadata(value);
  requiredText(value.fixtures.url, 'fixtures.url');
  requiredText(value.fixtures.metadata, 'fixtures.metadata');
  if (!Number.isInteger(value.fixtures.count) || value.fixtures.count < 1) throw new Error('Invalid fixtures.count');
  if (!Array.isArray(value.tasks) || !value.tasks.length) throw new Error('Missing tasks');
  for (const task of value.tasks) {
    for (const field of ['id', 'format', 'title', 'prompt', 'fixture'] as const) requiredText(task?.[field], 'tasks.' + field);
    if (!Number.isInteger(task.bytes) || task.bytes < 1) throw new Error('Missing or invalid tasks.bytes');
    if (typeof task.sha256 !== 'string' || !/^[a-f0-9]{64}$/.test(task.sha256)) throw new Error('Missing or invalid tasks.sha256');
    if (task.tier !== undefined && task.tier !== 'basic' && task.tier !== 'complex') throw new Error('Invalid tasks.tier');
  }
  if (!Array.isArray(value.runs) || !value.runs.length) throw new Error('Missing runs');
  if (new Set(value.arms.map(a => a.id)).size !== value.arms.length || new Set(value.tasks.map(t => t.id)).size !== value.tasks.length) throw new Error('Duplicate task or arm definition');
  const identities = new Set<string>();
  for (const run of value.runs) {
    if (!value.tasks.some(t => t.id === run.task) || !value.arms.some(a => a.id === run.arm)) throw new Error('Unknown run task or arm');
    requiredText(run.model, 'runs.model');
    if (!validDate(run.date) || !Number.isInteger(run.run) || run.run < 1 || typeof run.pass !== 'boolean') throw new Error('Invalid run');
    if (identities.has(key(run))) throw new Error('Duplicate run');
    identities.add(key(run));
  }
}

export function validateSavedOutput(result: Report['results'][number], check: { exitCode: number; stdout: string; stderr: string }) {
  const message = check.stdout.trim();
  const prefix = `${check.exitCode === 0 ? 'PASS' : 'FAIL'}: ${result.task}: `;
  if (![0, 1].includes(check.exitCode) || !message.startsWith(prefix)) throw new Error('Checker could not score ' + result.task + ': ' + check.stderr);
  const pass = check.exitCode === 0;
  if (pass && !result.pass) throw new Error('Runner rejected a workflow whose saved output now passes: ' + result.task);
  if (!pass && result.pass) throw new Error('Runner accepted a workflow whose saved output now fails: ' + result.task);
  return { pass, check_message: message.slice(prefix.length) };
}
const order = (a: string, b: string) => a < b ? -1 : a > b ? 1 : 0;
const key = (r: Pick<Run, 'model' | 'arm' | 'task' | 'run'>) => JSON.stringify([r.model, r.arm, r.task, r.run]);
const validDate = (date: string) => /^\d{4}-\d{2}-\d{2}$/.test(date) && new Date(date).toISOString().slice(0, 10) === date;
const invalidCause = (r: Report['results'][number]): Run['cause'] => r.timed_out ? 'timeout'
  : r.agent_exit_code !== undefined && r.agent_exit_code !== 0 ? 'agent-exit'
  : r.tool_calls === 0 ? 'no-tool-calls' : null;

export function savedOutputPath(directory: string, task: SourceTask, report: Report, result: Report['results'][number], suite: string, fallbackRun = 1) {
  const parts = [directory, task.id];
  if (report.run_directories ?? suite === 'v3') parts.push(`run-${result.run ?? fallbackRun}`);
  parts.push(task.tier === 'complex' && task.expected && 'answer' in task.expected ? 'answer.json' : task.output);
  return resolve(...parts);
}

export function parseArgs(args: string[]) {
  const groups: Group[] = [];
  let suite = '', reasons = '', output = '', meta = '';
  let arms = '', fixtures = '', tasks = '', keepInvalid = false;
  for (let i = 0; i < args.length; i++) {
    const flag = args[i];
    if (flag === '--keep-invalid') { keepInvalid = true; continue; }
    const value = args[++i];
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
    else if (flag === '--meta') meta = value;
    else if (flag === '--tasks') tasks = value;
    else if (flag === '--fixtures') fixtures = value;
    else throw new Error(`Unknown option ${flag}`);
  }
  const source = suite === 'v3' ? resolve(import.meta.dir, 'v3') : import.meta.dir;
  fixtures ||= resolve(source, 'fixtures.json');
  tasks ||= resolve(source, 'tasks.json');
  if (!suite || !groups.length) throw new Error('Required: --suite and --runs with --date per group');
  for (const g of groups) {
    if (!validDate(g.date) || !Number.isInteger(g.run) || g.run < 1) throw new Error('Invalid group date or run index');
  }
  if (!meta) throw new Error('Required: --meta <suite metadata JSON>');
  return { groups, suite, reasons, output, arms, fixtures, meta, tasks, keepInvalid };
}

export function collectResults(options: {
  suite: string; tasks: Task[]; arms: Arm[]; fixtures: Suite['fixtures']; reasons: Reason[]; keepInvalid?: boolean;
  groups: (Group & { report: Report })[]; harness: string; limits: string[]; reproduce: string[]; checker_validation?: Suite['checker_validation'];
}): Suite {
  validateMetadata(options);
  const tasks = new Map(options.tasks.map(t => [t.id, t]));
  const arms = new Map(options.arms.map(a => [a.id, a]));
  if (tasks.size !== options.tasks.length || arms.size !== options.arms.length) throw new Error('Duplicate task or arm definition');
  const reasons = new Map<string, Reason>();
  for (const r of options.reasons) {
    if (!tasks.has(r.task) || !arms.has(r.arm)) throw new Error(`Unknown reason task or arm: ${r.task}/${r.arm}`);
    if (reasons.has(key(r))) throw new Error(`Duplicate reason: ${key(r)}`);
    if (!r.reason?.trim() || !['library', 'model', 'timeout', 'agent-exit', 'no-tool-calls'].includes(r.cause ?? '')) throw new Error(`Invalid failure reason: ${r.task}`);
    reasons.set(key(r), r);
  }
  const runs: Run[] = [], seen = new Set<string>(), used = new Set<string>();
  for (const g of options.groups) {
    if (!arms.has(g.arm)) throw new Error(`Unknown arm: ${g.arm}; supply --arms for additional arms`);
    if (!validDate(g.date) || !Number.isInteger(g.run) || g.run < 1 || !g.report.model) throw new Error('Invalid run group');
    for (const r of g.report.results) {
      if (!tasks.has(r.task)) throw new Error(`Unknown task: ${r.task}`);
      if (typeof r.pass !== 'boolean' || typeof r.check_message !== 'string' || !Number.isInteger(r.tool_calls) || r.tool_calls < 0 || !Number.isFinite(r.duration_seconds) || r.duration_seconds < 0) throw new Error(`Invalid result: ${r.task}`);
      const run = r.run === undefined ? g.run : r.run;
      if (!Number.isInteger(run) || run < 1) throw new Error(`Invalid run index: ${r.task}`);
      const cause = invalidCause(r);
      if (cause && (!options.keepInvalid || r.pass)) throw new Error('Runner did not complete a valid workflow: ' + r.task);
      const identity = { model: g.report.model, arm: g.arm, task: r.task, run };
      const id = key(identity), reason = reasons.get(id);
      if (seen.has(id)) throw new Error(`Duplicate run: ${id}`);
      seen.add(id); used.add(r.task);
      if (!r.pass && options.suite === 'v2' && !reason && !cause) throw new Error(`Failed v2 run without a reason: ${id}`);
      if (!r.pass && (cause || options.suite === 'v3' && !reason) && !r.check_message.trim()) throw new Error(`Missing failure reason: ${id}`);
      runs.push({ ...identity, date: g.date, pass: r.pass, check: r.check_message,
        reason: r.pass ? null : cause ? r.check_message : reason?.reason ?? (options.suite === 'v3' ? r.check_message : null),
        cause: r.pass ? null : cause ?? reason?.cause ?? null,
        tool_calls: r.tool_calls, seconds: r.duration_seconds });
    }
  }
  if (!runs.length) throw new Error('No results');
  runs.sort((a, b) => order(a.model, b.model) || options.arms.findIndex(x => x.id === a.arm) - options.arms.findIndex(x => x.id === b.arm) || order(a.task, b.task) || a.run - b.run);
  const result: Suite = { schema: 1, suite: options.suite, date: runs.map(r => r.date).sort().at(-1)!,
    fixtures: options.fixtures, harness: options.harness, limits: options.limits, reproduce: options.reproduce,
    arms: options.arms.filter(a => runs.some(r => r.arm === a.id)),
    tasks: [...used].sort(order).map(id => {
      const t = tasks.get(id)!;
      return { id: t.id, ...(t.tier !== undefined ? { tier: t.tier } : {}), format: t.format, title: t.title, prompt: t.prompt, fixture: t.fixture, bytes: t.bytes, sha256: t.sha256 };
    }), runs, ...(options.checker_validation ? { checker_validation: options.checker_validation } : {}) };
  validateSuite(result);
  return result;
}

export const serialize = (suite: Suite) => JSON.stringify(suite, null, 2) + '\n';

export async function main(args: string[]) {
  const options = parseArgs(args);
  let sourceTasks: SourceTask[] = JSON.parse(await readFile(options.tasks, 'utf8'));
  if (options.suite === 'v3') {
    const complex: SourceTask[] = JSON.parse(await readFile(resolve(import.meta.dir, 'v3/complex.json'), 'utf8'));
    sourceTasks = [...sourceTasks.map(t => ({ ...t, tier: 'basic' as const })), ...complex.map(t => ({ ...t, tier: 'complex' as const }))];
  }
  const manifest = JSON.parse(await readFile(options.fixtures, 'utf8'));
  const metadata: SuiteMetadata = JSON.parse(await readFile(options.meta, 'utf8'));
  validateMetadata(metadata);
  const tasks = sourceTasks.map(task => {
    const pin = manifest.fixtures[task.fixture];
    if (!pin) throw new Error('Missing fixture pin: ' + task.fixture);
    return { ...task, bytes: pin.bytes as number, sha256: pin.sha256 as string };
  });
  const arms: Arm[] = options.arms ? JSON.parse(await readFile(options.arms, 'utf8')) : metadata.arms;
  const reasons: Reason[] = options.reasons ? JSON.parse(await readFile(options.reasons, 'utf8')) : [];
  const groups = [];
  const environment = { ...process.env };
  environment.FORMAT_EVAL_TASKS = resolve(options.tasks);
  environment.FORMAT_EVAL_MANIFEST = resolve(options.fixtures);
  for (const g of options.groups) {
    const report: Report = JSON.parse(await readFile(resolve(g.directory, 'results.json'), 'utf8'));
    for (const r of report.results) {
      const task = tasks.find(t => t.id === r.task);
      if (!task) throw new Error(`Unknown task: ${r.task}`);
      if (r.run !== undefined && (!Number.isInteger(r.run) || r.run < 1)) throw new Error(`Invalid run index: ${r.task}`);
      if (invalidCause(r)) {
        if (!options.keepInvalid || r.pass) throw new Error('Runner did not complete a valid workflow: ' + r.task);
        continue;
      }
      const output = savedOutputPath(g.directory, task, report, r, options.suite, g.run);
      const check = Bun.spawnSync(['python3', resolve(import.meta.dir, 'check.py'), ...(options.suite === 'v3' ? ['--suite', 'v3'] : []), r.task, output], {
        env: environment,
      });
      Object.assign(r, validateSavedOutput(r, { exitCode: check.exitCode, stdout: check.stdout.toString(), stderr: check.stderr.toString() }));
    }
    groups.push({ ...g, report });
  }
  const origin = manifest.origin.replace(/\/$/, '');
  const result = collectResults({ ...options, ...metadata, tasks, arms, reasons, groups,
    fixtures: { kind: metadata.fixtures.kind, count: new Set(groups.flatMap(g => g.report.results.map(r => tasks.find(t => t.id === r.task)!.fixture))).size,
      url: origin + '/', metadata: origin + '/metadata.json', description: metadata.fixtures.description } });
  if (options.output) await writeFile(options.output, serialize(result));
  else process.stdout.write(serialize(result));
}

if (import.meta.main) await main(process.argv.slice(2));
