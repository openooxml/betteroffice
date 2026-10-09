import { FORMATS, type Format } from './summary';

export const FRAME_MS = 1000 / 60;

export interface OpGroup {
  name: string;
  count: number;
  p50: number;
  p95: number;
  max: number;
}

export interface Case {
  scenario: string;
  sample: string;
  status: string;
  description: string;
  participants: string[];
  opCount: number;
  totalMs: number;
}

export interface FormatLatency {
  format: Format;
  cases: Case[];
  passed: number;
  ops: number;
  errors: number;
  p50: number;
  p95: number;
  withinFrame: number;
  groups: OpGroup[];
  values: number[];
}

export interface Environment {
  cpu: string;
  cpus: number;
  platform: string;
  arch: string;
}

export interface Latency {
  commit: string;
  recordedAt: string | null;
  environment: Environment | null;
  formats: Partial<Record<Format, FormatLatency>>;
  cases: number;
  passed: number;
  ops: number;
  p50: number;
  p95: number;
  withinFrame: number;
  multiEditor: number;
  crossSdk: number;
}

export interface LatestE2e {
  sha: string;
  publishedAt: string | null;
  formats: Format[];
}

type Json = Record<string, unknown>;

const object = (value: unknown): Json | null =>
  value && typeof value === 'object' && !Array.isArray(value) ? (value as Json) : null;
const list = (value: unknown): Json[] =>
  Array.isArray(value) ? value.map(object).filter((entry): entry is Json => entry !== null) : [];
const finite = (value: unknown): number | null =>
  typeof value === 'number' && Number.isFinite(value) && value >= 0 ? value : null;

/** Nearest-rank percentile, the same rule the e2e harness reports. */
export function percentile(sorted: number[], p: number): number {
  if (!sorted.length) return 0;
  const index = Math.min(sorted.length - 1, Math.ceil((p / 100) * sorted.length) - 1);
  return sorted[Math.max(0, index)]!;
}

/** Operation names are `verb:variant`; variants of one verb share a group. */
export const groupName = (op: string): string => op.split(':')[0] || op;

function formatLatency(run: Json, format: Format): FormatLatency {
  const cases: Case[] = [];
  const values: number[] = [];
  const grouped = new Map<string, number[]>();
  let errors = 0;
  for (const scenario of list(run.scenarios)) {
    const ops = list(scenario.ops);
    const participants = Array.isArray(scenario.participants) ? scenario.participants.map(String) : [];
    let totalMs = 0;
    for (const op of ops) {
      const ms = finite(op.e2eMs);
      if (ms === null) continue;
      if (typeof op.error === 'string') errors += 1;
      totalMs += ms;
      values.push(ms);
      const name = groupName(String(op.op));
      const group = grouped.get(name) ?? [];
      if (!group.length) grouped.set(name, group);
      group.push(ms);
    }
    cases.push({
      scenario: String(scenario.scenario),
      sample: String(scenario.sample),
      status: String(scenario.status),
      description: typeof scenario.description === 'string' ? scenario.description : '',
      participants,
      opCount: ops.length,
      totalMs,
    });
  }
  values.sort((left, right) => left - right);
  const groups = [...grouped].map(([name, samples]) => {
    samples.sort((left, right) => left - right);
    return {
      name,
      count: samples.length,
      p50: percentile(samples, 50),
      p95: percentile(samples, 95),
      max: samples.at(-1) ?? 0,
    };
  });
  return {
    format,
    cases,
    passed: cases.filter((entry) => entry.status === 'passed').length,
    ops: values.length,
    errors,
    p50: percentile(values, 50),
    p95: percentile(values, 95),
    withinFrame: values.filter((ms) => ms <= FRAME_MS).length,
    groups: groups.sort((left, right) => right.count - left.count || left.name.localeCompare(right.name)),
    values,
  };
}

/** Aggregates the per-format result files one end-to-end run recorded. */
export function summarizeLatency(runs: unknown[]): Latency {
  const formats: Partial<Record<Format, FormatLatency>> = {};
  let commit = '';
  let recordedAt: string | null = null;
  let environment: Environment | null = null;
  for (const value of runs) {
    const run = object(value);
    const format = FORMATS.find((entry) => entry === run?.format);
    if (!run || run.schemaVersion !== 3 || !format || formats[format]) throw new Error('Malformed e2e results');
    if (!/^[a-f0-9]{40}$/.test(String(run.commit)) || (commit && run.commit !== commit))
      throw new Error('e2e results disagree on the commit');
    commit = String(run.commit);
    if (typeof run.recordedAt === 'string' && (!recordedAt || run.recordedAt > recordedAt))
      recordedAt = run.recordedAt;
    const machine = object(run.environment);
    if (machine && !environment)
      environment = {
        cpu: String(machine.cpu ?? 'unknown'),
        cpus: Number(machine.cpus ?? 0),
        platform: String(machine.platform ?? 'unknown'),
        arch: String(machine.arch ?? 'unknown'),
      };
    formats[format] = formatLatency(run, format);
  }
  if (!commit) throw new Error('Malformed e2e results');
  const all = Object.values(formats);
  const cases = all.flatMap((entry) => entry.cases);
  const pooled = all.flatMap((entry) => entry.values).sort((left, right) => left - right);
  return {
    commit,
    recordedAt,
    environment,
    formats,
    cases: cases.length,
    passed: cases.filter((entry) => entry.status === 'passed').length,
    ops: pooled.length,
    p50: percentile(pooled, 50),
    p95: percentile(pooled, 95),
    withinFrame: all.reduce((sum, entry) => sum + entry.withinFrame, 0),
    multiEditor: cases.filter((entry) => entry.participants.length > 1).length,
    crossSdk: cases.filter((entry) => entry.participants.includes('python')).length,
  };
}

export function parseLatestE2e(value: unknown): LatestE2e {
  const root = object(value);
  if (!root || !/^[a-f0-9]{40}$/.test(String(root.sha))) throw new Error('Malformed e2e manifest');
  const formats = Array.isArray(root.formats)
    ? FORMATS.filter((format) => (root.formats as unknown[]).includes(format))
    : [];
  if (!formats.length) throw new Error('Malformed e2e manifest');
  return {
    sha: String(root.sha),
    publishedAt: typeof root.published_at_utc === 'string' ? root.published_at_utc : null,
    formats,
  };
}

export const e2eUrl = (sha: string, format: Format, base = '/e2e'): string => {
  if (!/^[a-f0-9]{40}$/.test(sha)) throw new Error('Expected a full commit SHA');
  return `${base}/${sha}/${format}.json`;
};
